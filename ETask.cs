using System;
using System.Runtime.CompilerServices;
using System.Windows.Threading;

namespace ExcelAddIn1
{
    public class ETask
    {
        /// <summary>
        /// 非同期で条件が満たされるのを待つ
        /// </summary>
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
        readonly Func<bool> _predicate;
        //引数: 条件
        public WaitUntilAsync(Func<bool> predicate) { _predicate = predicate; }
        public WaitUntilAsync GetAwaiter() { return this; }
        //falseにしないとすぐにGetResult→終了する。OnCompletedが呼ばれない。
        public bool IsCompleted => false;
        //continuationを呼ぶとawaitを抜ける
        public void OnCompleted(System.Action continuation) => AsyncSupport.Add(new FuncCore(continuation, _predicate)); 
        //最後に呼ばれる。
        public void GetResult() { }
    }
    
    /// <summary>
    /// Asyncのカスタム用
    /// </summary>
    readonly struct AsyncSupport
    {
        static FuncCore[] Core;
        static int Count;
        static object lockg;
        static MyTimer timer;
        static AsyncSupport()
        {
            Core = new FuncCore[4];
            Count = default(int);
            lockg = new object();
            timer = new MyTimer(16);
        }
        /// <summary>
        /// OnCompleted関数専用
        /// </summary>
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
                timer.Automatic();
            }
        }
        /// <summary>
        /// Countが0以外の場合は常に呼ばれる
        /// </summary>
        static void RunCore(object sender, EventArgs e)
        {
            lock (lockg)
            {
                for (int i = 0; i < Count; i++)
                {
                    FuncCore core = Core[i];
                    if (core.Equals(null)) continue;
                    //判定
                    if (core.predicate == null || core.predicate())
                    {
                        Count--;
                        Core[i] = Core[Count];
                        Core[Count] = default(FuncCore);
                        timer.Automatic();
                        //待機終了
                        core.core();
                    }
                }
            }
        }
        class MyTimer
        {
            DispatcherTimer timer;
            public MyTimer(int Interval)
            {
                timer = new DispatcherTimer();
                timer.Tick += new EventHandler(RunCore);
                timer.Interval = new TimeSpan(0, 0, 0, 0, Interval);
                timer.IsEnabled = false;
            }
            /// <summary>
            /// タイマーの起動、停止を自動で行う。
            /// </summary>
            public void Automatic()
            {
                if (Count == 0) timer.Stop();
                if (Count == 1) timer.Start();
                System.Threading.Thread.Yield();
            }
        }
    }
    //定義
    readonly struct FuncCore
    {
        public readonly System.Action core;
        public readonly Func<bool> predicate;
        public FuncCore(in System.Action core, in Func<bool> predicate) 
        { this.core = core; this.predicate = predicate; }
    }
}