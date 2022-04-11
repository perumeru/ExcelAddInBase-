using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public class MyTask
    {
        //条件を待つ
        public static WaitUntilAsync WaitUntil(in Func<bool> predicate)
        {
            return new WaitUntilAsync(predicate);
        }
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
        public void OnCompleted(System.Action continuation) => AsyncSupport.Add(new FuncCore(continuation, _predicate)); 
        //最後に呼ばれる。
        public void GetResult() { }
    }
    /// <summary>
    /// Asyncのカスタム用
    /// </summary>
    public readonly struct AsyncSupport
    {
        static FuncCore[] Core = new FuncCore[4];
        static uint Count = default(uint);
        static object lockg = new object();
        public static void Add(in FuncCore core)
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
        static void RunCore(object sender, EventArgs e)
        {
            if (Count == 0) return;
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
        //タイマー
        class MyTimer
        {
            const int Interval = 20;
            Timer timer = new Timer();
            public MyTimer()
            {
                timer.Tick += new EventHandler(RunCore);
                timer.Interval = Interval;
                timer.Enabled = true;
            }
            ~MyTimer()
            {
                timer.Dispose();
            }
        }
        static AsyncSupport() => new MyTimer();
    }

    //定義
    public readonly struct FuncCore
    {
        public readonly System.Action core;
        public readonly Func<bool> predicate;
        public FuncCore(in System.Action core, in Func<bool> predicate) { this.core = core; this.predicate = predicate; }
    }
}