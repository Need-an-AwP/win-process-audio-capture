#include <node.h>
#include <v8.h>
#include <cstdlib>
#include <stdlib.h>
#include <time.h>



class MyClass
{
public:
    MyClass()
    {
        srand(time(NULL));
    }

    static void GetRandomNumber(const v8::FunctionCallbackInfo<v8::Value> &args)
    {
        v8::Isolate *isolate = args.GetIsolate();

        // 创建一个新的 MyClass 实例
        MyClass *object = new MyClass();

        // 在异步回调中获取随机数并返回给 JavaScript
        object->GetRandomNumberAsync(args);

        // 释放 MyClass 实例
        delete object;
    }

private:
    void GetRandomNumberAsync(const v8::FunctionCallbackInfo<v8::Value> &args)
    {
        v8::Isolate *isolate = args.GetIsolate();

        // 模拟异步操作，比如定时器或者事件触发等方式
        // 这里使用了简单的延迟以模拟异步操作
        double randomNumber = static_cast<double>(rand()) / RAND_MAX;

        // 返回随机数给 JavaScript
        args.GetReturnValue().Set(v8::Number::New(isolate, randomNumber));
    }
};

std::function<void(int)> callback;
void Method(const v8::FunctionCallbackInfo<v8::Value> &args)
{
    v8::Isolate *isolate = args.GetIsolate();
    // args.GetReturnValue().Set(v8::String::NewFromUtf8(isolate, "world").ToLocalChecked());
}

void init(v8::Local<v8::Object> exports)
{
    NODE_SET_METHOD(exports, "exports", Method);
}


#undef NODE_MODULE_VERSION
#define NODE_MODULE_VERSION 121
NODE_MODULE(NODE_GYP_MODULE_NAME, init)