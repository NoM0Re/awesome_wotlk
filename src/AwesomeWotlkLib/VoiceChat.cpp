#include "VoiceChat.h"
#include "GameClient.h"
#include "Hooks.h"
#include <sapi.h>
#include <sphelper.h>
#include <codecvt>
#include <locale>
#include <vector>
#include <string>
#include <Windows.h>

static Console::CVar* s_cvar_voiceID;
static Console::CVar* s_cvar_speed;
static Console::CVar* s_cvar_volume;

// Globale SAPI-Voice-Instanz
static ISpVoice* g_pVoice = nullptr;

std::string WideStringToUtf8(const std::wstring& wstr)
{
    if (wstr.empty())
        return std::string();
    int sizeNeeded = WideCharToMultiByte(
        CP_UTF8, 0, wstr.c_str(), -1,
        nullptr, 0, nullptr, nullptr
    );
    if (sizeNeeded == 0) {
        return std::string();
    }
    std::string result(sizeNeeded - 1, '\0');
    WideCharToMultiByte(
        CP_UTF8, 0, wstr.c_str(), -1,
        &result[0], sizeNeeded - 1, nullptr, nullptr
    );
    return result;
}

std::wstring Utf8ToWide(const char* utf8Str)
{
    if (!utf8Str) return L"";
    int size = MultiByteToWideChar(CP_UTF8, 0, utf8Str, -1, nullptr, 0);
    if (size == 0) return L"";
    std::wstring wide(size - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8Str, -1, &wide[0], size);
    return wide;
}

struct VoiceTtsVoiceType {
    int voiceID;
    std::wstring name;
};

// ====================
// COM / Voice Init (only once)
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

static void VoiceChat_InitVoice()
{
    if (!g_pVoice) {
        VoiceChat_InitCOM();
        HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&g_pVoice);
        if (FAILED(hr)) {
            g_pVoice = nullptr;
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

// 'destination' parameter is reserved for future use (e.g., remote TTS output)
static void VoiceChat_SpeakText(int voiceID, const std::wstring& text, const std::wstring& destination, int rate = 0, int volume = 100)
{
    VoiceChat_InitVoice();
    (void)destination; // Suppress unused parameter warning
    if (!g_pVoice) return;

    // Stimme setzen (falls vorhanden)
    IEnumSpObjectTokens* pEnum = nullptr;
    ULONG count = 0;
    HRESULT hr = SpEnumTokens(SPCAT_VOICES, NULL, NULL, &pEnum);

    if (SUCCEEDED(hr) && pEnum)
    {
        ISpObjectToken* pToken = nullptr;
        ULONG index = 0;
        while (pEnum->Next(1, &pToken, &count) == S_OK && count)
        {
            if (index == static_cast<ULONG>(voiceID))
            {
                g_pVoice->SetVoice(pToken);
                pToken->Release();
                break;
            }
            pToken->Release();
            ++index;
        }
        pEnum->Release();
    }

    g_pVoice->SetRate(rate);
    g_pVoice->SetVolume(volume);
    // ASYNC → kein Freeze, g_pVoice bleibt erhalten
    g_pVoice->Speak(text.c_str(), SPF_ASYNC, NULL);
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

        lua_pushnumber(L, voice.voiceID);
        lua_setfield(L, -2, "voiceID");

        std::string nameUtf8 = WideStringToUtf8(voice.name);
        lua_pushstring(L, nameUtf8.c_str());
        lua_setfield(L, -2, "name");

        lua_rawseti(L, -2, i++);
    }
    return 1;
}

static int Lua_VoiceChat_GetRemoteTtsVoices(lua_State* L)
{
    auto voices = VoiceChat_GetTtsVoices();
    lua_createtable(L, 0, 0);

    int i = 1;
    for (const auto& voice : voices)
    {
        lua_newtable(L);

        lua_pushnumber(L, voice.voiceID);
        lua_setfield(L, -2, "voiceID");

        std::string nameUtf8 = WideStringToUtf8(voice.name);
        lua_pushstring(L, nameUtf8.c_str());
        lua_setfield(L, -2, "name");

        lua_rawseti(L, -2, i++);
    }
    return 1;
}

static int OnVoiceIDChanged(Console::CVar* cvar, const char* prevVal, const char* newVal, void* udata)
{
    int val = atoi(newVal);
    if (val < 0) val = 0;
    std::string valStr = std::to_string(val);
    SetCVarValue(cvar, valStr.c_str(), 0, 0, 0, 0);
    return 1;
}

static int OnSpeedChanged(Console::CVar* cvar, const char* prevVal, const char* newVal, void* udata)
{
    int val = atoi(newVal);
    if (val < -10) val = -10;
    else if (val > 10) val = 10;
    std::string valStr = std::to_string(val);
    SetCVarValue(cvar, valStr.c_str(), 0, 0, 0, 0);
    return 1;
}

static int OnVolumeChanged(Console::CVar* cvar, const char* prevVal, const char* newVal, void* udata)
{
    int val = atoi(newVal);
    if (val < 0) val = 0;
    else if (val > 100) val = 100;
    std::string valStr = std::to_string(val);
    SetCVarValue(cvar, valStr.c_str(), 0, 0, 0, 0);
    return 1;
}

static int Lua_VoiceChat_SpeakText(lua_State* L)
{
    int voiceID = (int)luaL_checknumber(L, 1);
    const char* text = luaL_checklstring(L, 2, nullptr);

    const char* destination = "default";
    if (lua_gettop(L) >= 3 && lua_type(L, 3) != LUA_TNIL)
        destination = luaL_checklstring(L, 3, nullptr);

    int rate = 0;
    if (lua_gettop(L) >= 4 && lua_type(L, 4) != LUA_TNIL)
        rate = (int)luaL_checknumber(L, 4);

    int volume = 100;
    if (lua_gettop(L) >= 5 && lua_type(L, 5) != LUA_TNIL)
        volume = (int)luaL_checknumber(L, 5);

    std::wstring wText = Utf8ToWide(text);
    std::wstring wDest = Utf8ToWide(destination);

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
    Hooks::FrameXML::registerCVar(&s_cvar_voiceID, "tts_voiceID", NULL, (Console::CVarFlags)1, "0", OnVoiceIDChanged);
    Hooks::FrameXML::registerCVar(&s_cvar_speed, "tts_speed", NULL, (Console::CVarFlags)1, "0", OnSpeedChanged);
    Hooks::FrameXML::registerCVar(&s_cvar_volume, "tts_volume", NULL, (Console::CVarFlags)1, "100", OnVolumeChanged);
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

void VoiceChat::initialize()
{
    VoiceChat_InitVoice(); // erstellt g_pVoice einmalig
    Hooks::FrameXML::registerLuaLib(lua_openlibvoicechat);
    RegisterVoiceChatCVars();
}
