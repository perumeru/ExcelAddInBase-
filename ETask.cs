using System;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace ExcelAddIn1
{
    //定義
    public readonly struct FuncCore
    {
        public readonly System.Action core;
        public readonly Func<bool> predicate;
        public FuncCore(in System.Action core, in Func<bool> predicate) { this.core = core; this.predicate = predicate; }
    }
    /// <summary>
    /// 非同期で条件が満たされるのを待つ
    /// </summary>
    public struct WaitUntilAsync : INotifyCompletion
    {
        Func<bool> _predicate;
        //引数: 条件
        public WaitUntilAsync(Func<bool> predicate) { _predicate = predicate; }
        public WaitUntilAsync GetAwaiter() { return this; }
        //falseにしないとすぐにGetResult→終了する。OnCompletedが呼ばれない。
        public bool IsCompleted => false;
        //continuationを呼ぶと非同期が終了する。
        public void OnCompleted(System.Action continuation) { 
            ETask.asyncSupport.Add(new FuncCore(continuation, _predicate)); 
        }
        //最後に呼ばれる。
        public void GetResult() { }
    }
    public class MyTask
    {
        //スマートにする為に
        public static WaitUntilAsync WaitUntil(in Func<bool> predicate)
        {
            return new WaitUntilAsync(predicate);
        }
    }
    /// <summary>
    /// Asyncのカスタム用
    /// </summary>
    public class ETask
    {
        const int Delay = 20;
        public readonly static AsyncSupport asyncSupport = new AsyncSupport(4);
        public class AsyncSupport
        {
            FuncCore[] Core;
            uint Count;
            //配列のアクセスが変にならないよう
            static object lockg = new object();
            public void Add(in FuncCore core)
            {
                lock (lockg)
                {
                    if (Core.Length <= Count)
                    {
                        FuncCore[] array2 = Core;
                        Core = new FuncCore[Core.Length * 2];
                        Array.Copy(array2, Core, Count);
                    }
                    Core[Count] = core;
                    Count++;
                }
            }
            //sleepかdelay位しかデフォルトで待てそうな処理がない?
            async Task RunCore()
            {
                while (Core.Length > 0)
                {
                    await Task.Delay(Delay);
                    lock (lockg)
                    {
                        for (int i = 0; i < Count; i++)
                        {
                            FuncCore core = Core[i];
                            if (core.Equals(null)) continue;
                            if (core.predicate == null || core.predicate())
                            {
                                //MyTask非同期が終了する。
                                core.core();
                                Count--;
                                Core[i] = Core[Count];
                                Core[Count] = default(FuncCore);
                            }
                        }
                    }
                }
            }
            public AsyncSupport(in uint arraysize)
            {
                Count = default(uint);
                Core = new FuncCore[arraysize];
                Task.Run(RunCore);
            }
        }
    }
}