#define NAPI_CPP_EXCEPTIONS 1
#include <node_api.h>
#include <random>
#include <napi.h>
#include <sstream>

Napi::Value DefaultCaputure(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();

    while (true)
    {
        return Napi::String::New(env, "asdasd");
    }
}

Napi::Object Init(Napi::Env env, Napi::Object exports)
{
    exports.Set(Napi::String::New(env, "defaultCaputure"), Napi::Function::New(env, DefaultCaputure));

    return exports;
}

NODE_API_MODULE(myncpp1, Init)