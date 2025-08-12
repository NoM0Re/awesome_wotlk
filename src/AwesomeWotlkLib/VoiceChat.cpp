#ifdef _WIN32
#include "VoiceChat.h"
#include "GameClient.h"
#include "Hooks.h"
#include <sapi.h>
#include <sphelper.h>
#include <codecvt>
#include <locale>
#include <vector>
#include <string>

CVar* s_cvar_voiceID = nullptr;
CVar* s_cvar_speed = nullptr;
CVar* s_cvar_volume = nullptr;

struct VoiceTtsVoiceType {
    int voiceID;
    std::wstring name;
};

// ====================
// COM Init (only once)
// ====================
static void VoiceChat_InitCOM()
{
    static bool comInitDone = false;
    if (!comInitDone) {
        HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
        if (SUCCEEDED(hr)) {
            comInitDone = true;
        }
    }
}

// ====================
// Static Voice Functions
// ====================

static std::vector<VoiceTtsVoiceType> VoiceChat_GetTtsVoices()
{
    VoiceChat_InitCOM();

    std::vector<VoiceTtsVoiceType> voices;
    IEnumSpObjectTokens* pEnum = nullptr;
    ULONG count = 0;

    HRESULT hr = SpEnumTokens(SPCAT_VOICES, NULL, NULL, &pEnum);
    if (SUCCEEDED(hr) && pEnum)
    {
        ISpObjectToken* pToken = nullptr;
        ULONG index = 0;
        while (pEnum->Next(1, &pToken, &count) == S_OK && count)
        {
            WCHAR* description = nullptr;
            if (SUCCEEDED(SpGetDescription(pToken, &description)) && description)
            {
                voices.push_back({ static_cast<int>(index), description });
                CoTaskMemFree(description);
            }
            pToken->Release();
            ++index;
        }
        pEnum->Release();
    }

    return voices;
}

static std::vector<VoiceTtsVoiceType> VoiceChat_GetRemoteTtsVoices()
{
    return VoiceChat_GetTtsVoices();
}

static void VoiceChat_SpeakText(int voiceID, const std::wstring& text, const std::wstring& /*destination*/, int rate = 0, int volume = 100)
{
    VoiceChat_InitCOM();

    ISpVoice* pVoice = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);
    if (FAILED(hr) || !pVoice)
        return;

    IEnumSpObjectTokens* pEnum = nullptr;
    ULONG count = 0;
    hr = SpEnumTokens(SPCAT_VOICES, NULL, NULL, &pEnum);

    if (SUCCEEDED(hr) && pEnum)
    {
        ISpObjectToken* pToken = nullptr;
        ULONG index = 0;
        while (pEnum->Next(1, &pToken, &count) == S_OK && count)
        {
            if (index == static_cast<ULONG>(voiceID))
            {
                pVoice->SetVoice(pToken);
                pToken->Release();
                break;
            }
            pToken->Release();
            ++index;
        }
        pEnum->Release();
    }

    pVoice->SetRate(rate);
    pVoice->SetVolume(volume);
    pVoice->Speak(text.c_str(), SPF_ASYNC | SPF_IS_XML, NULL); // SPF_ASYNC = running parralel

    pVoice->Release();
}

// ====================
// Lua Bindings
// ====================

static int Lua_VoiceChat_GetTtsVoices(lua_State* L)
{
    auto voices = VoiceChat_GetTtsVoices();
    lua_createtable(L, 0, 0);

    int i = 1;
    for (const auto& voice : voices)
    {
        lua_newtable(L);
        lua_pushstring(L, "voiceID"); lua_pushinteger(L, voice.voiceID); lua_settable(L, -3);
        lua_pushstring(L, "name");
        std::string nameUtf8 = std::wstring_convert<std::codecvt_utf8<wchar_t>>().to_bytes(voice.name);
        lua_pushstring(L, nameUtf8.c_str()); lua_settable(L, -3);
        lua_rawseti(L, -2, i++);
    }
    return 1;
}

static int Lua_VoiceChat_GetRemoteTtsVoices(lua_State* L)
{
    auto voices = VoiceChat_GetRemoteTtsVoices();
    lua_createtable(L, 0, 0);

    int i = 1;
    for (const auto& voice : voices)
    {
        lua_newtable(L);
        lua_pushstring(L, "voiceID"); lua_pushinteger(L, voice.voiceID); lua_settable(L, -3);
        lua_pushstring(L, "name");
        std::string nameUtf8 = std::wstring_convert<std::codecvt_utf8<wchar_t>>().to_bytes(voice.name);
        lua_pushstring(L, nameUtf8.c_str()); lua_settable(L, -3);
        lua_rawseti(L, -2, i++);
    }
    return 1;
}

int OnVoiceIDChanged(CVar* cvar, const char* newVal)
{
    int val = atoi(newVal);
    if (val < 0) val = 0;
    std::string valStr = std::to_string(val);
    SetCVarValue(cvar, valStr.c_str(), 0, 0, 0, 0);
    return 0;
}

int OnSpeedChanged(CVar* cvar, const char* newVal)
{
    int val = atoi(newVal);
    if (val < -10) val = -10;
    else if (val > 10) val = 10;
    std::string valStr = std::to_string(val);
    SetCVarValue(cvar, valStr.c_str(), 0, 0, 0, 0);
    return 0;
}

int OnVolumeChanged(CVar* cvar, const char* newVal)
{
    int val = atoi(newVal);
    if (val < 0) val = 0;
    else if (val > 100) val = 100;
    std::string valStr = std::to_string(val);
    SetCVarValue(cvar, valStr.c_str(), 0, 0, 0, 0);
    return 0;
}

static int Lua_VoiceChat_SpeakText(lua_State* L)
{
    int voiceID = luaL_checkinteger(L, 1);
    const char* text = luaL_checkstring(L, 2);
    const char* destination = luaL_optstring(L, 3, "default");
    int rate = luaL_optinteger(L, 4, 0);
    int volume = luaL_optinteger(L, 5, 100);

    std::wstring wText = std::wstring_convert<std::codecvt_utf8<wchar_t>>().from_bytes(text);
    std::wstring wDest = std::wstring_convert<std::codecvt_utf8<wchar_t>>().from_bytes(destination);

    VoiceChat_SpeakText(voiceID, wText, wDest, rate, volume);
    return 0;
}

void VoiceChat_SpeakText(const std::wstring& text)
{
    int voiceID = atoi(s_cvar_voiceID->vStr);
    int speed = atoi(s_cvar_speed->vStr);
    int volume = atoi(s_cvar_volume->vStr);

    VoiceChat_SpeakText(voiceID, text, L"default", speed, volume);
}

void RegisterVoiceChatCVars()
{
    s_cvar_voiceID = RegisterCVar("tts_voiceID", "Voice ID for TTS", 1, "0", OnVoiceIDChanged, 0, 0, 0, 0);
    s_cvar_speed   = RegisterCVar("tts_speed", "Speech-rate for TTS", 1, "0", OnSpeedChanged, 0, 0, 0, 0);
    s_cvar_volume  = RegisterCVar("tts_volume", "Volume for TTS", 1, "100", OnVolumeChanged, 0, 0, 0, 0);
}

static int lua_openlibvoicechat(lua_State* L)
{
    luaL_Reg methods[] = {
        {"GetTtsVoices", Lua_VoiceChat_GetTtsVoices},
        {"GetRemoteTtsVoices", Lua_VoiceChat_GetRemoteTtsVoices},
        {"SpeakText", Lua_VoiceChat_SpeakText},
    };

    lua_createtable(L, 0, std::size(methods));
    for (size_t i = 0; i < std::size(methods); i++) {
        lua_pushcfunction(L, methods[i].func);
        lua_setfield(L, -2, methods[i].name);
    }
    lua_setglobal(L, "C_VoiceChat");
    return 0;
}

// ====================
// Initialization
// ====================

namespace VoiceChat {
    void initialize()
    {
        VoiceChat_InitCOM();
        Hooks::FrameXML::registerLuaLib(lua_openlibvoicechat);
        RegisterVoiceChatCVars();
    }
}

#else // ifndef _WIN32

namespace VoiceChat {
    void initialize() {
        // No implementation for non-Windows platforms
    }
}

#endif
