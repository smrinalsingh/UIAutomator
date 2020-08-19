using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Windows.Automation;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Runtime.InteropServices;
using System;
using System.IO;

namespace UIAutomator
{
    [Cmdlet(VerbsCommon.Get, "uiWindow")]
    [OutputType(typeof(AutomationElement))]
    public class GetuiWindowCmdlet : Cmdlet
    {
        [Parameter(Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Mandatory = true)]
        public string Name { get; set; }
        [Parameter(Position = 1)]
        public string Class { get; set; }
        [Parameter(Position = 2)]
        public string AutoID { get; set; }
        [Parameter(Position = 3)]
        public int Timeout { get; set; } = 5000;
        [Parameter(Position = 4)]
        public int Index { get; set; } = 1;
        public AutomationElement ParentWin { get; set; }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
        }

        protected override void ProcessRecord()
        {
            AutomationElement _RootElem = AutomationElement.RootElement;
            ParentWin = StaticMethods.FindAtIndex(_RootElem, Name, Class, "window", AutoID, Timeout, Index);
            WriteObject(ParentWin);
        }
    }

    [Cmdlet(VerbsCommon.Get, "uiControl")]
    [OutputType(typeof(AutomationElement))]
    public class GetuiControlCmdlet : Cmdlet
    {
        [Parameter(Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Mandatory = true)]
        [Alias("Win")]
        public AutomationElement ParentWin { get; set; }

        [Parameter(Position = 1)]
        [Alias("N")]
        public string Name { get; set; }

        [Parameter(Position = 2)]
        [Alias("ElemType")]
        public string Type { get; set; }

        [Parameter(Position = 3)]
        [Alias("C")]
        public string Class { get; set; }

        [Parameter(Position = 4)]
        [Alias("ID")]
        public string AutoID { get; set; }

        [Parameter(Position = 5)]
        public int Index { get; set; } = 1;

        [Parameter(Position = 6)]
        [Alias("t")]
        public int Timeout { get; set; } = 5000;

        public AutomationElement UIElement { get; set; }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
        }
        protected override void ProcessRecord()
        {
            UIElement = null;
            if (ParentWin != null)
            {
                UIElement = StaticMethods.FindUnderElement(ParentWin, Name, Class, AutoID, Type, Timeout, Index);
            }
            WriteObject(UIElement);
        }
    }

    [Cmdlet(VerbsLifecycle.Invoke, "uiClick")]
    [OutputType(typeof(AutomationElement))]
    public class InvokeuiClickCmdlet : Cmdlet
    {
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public AutomationElement UIElement { get; set; }
        [Parameter(Position = 1)]
        public string MouseButton { get; set; } = "left";
        [Parameter(Position = 2)]
        [Alias("t")]
        public int Timeout { get; set; } = 5000;
        [Parameter(Position = 3)]
        public SwitchParameter Focus { get; set; } = false;

        protected override void ProcessRecord()
        {
            if (Focus.ToBool())
                UIElement.SetFocus();
            StaticMethods.ElementClick(UIElement, MouseButton);
            //WriteObject(UIElement);
        }
    }

    [Cmdlet(VerbsLifecycle.Invoke, "uiSendKeys")]
    [OutputType(typeof(AutomationElement))]
    public class InvokeuiSendKeysCmdlet : Cmdlet
    {
        [Parameter(Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public AutomationElement UIElement { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public string Text { get; set; }

        [Parameter(Position = 2)]
        public SwitchParameter Click { get; set; }

        [Parameter(Position = 3)]
        [Alias("t")]
        public int Timeout { get; set; } = 5000;

        protected override void ProcessRecord()
        {
            if (UIElement == null && Click.ToBool())
            {
                WriteError(new ErrorRecord(new ElementNotAvailableException("Element to be clicked before typing isn't available. Unable to execute."), "1", ErrorCategory.InvalidOperation, UIElement));
                return;
            }
            else if (UIElement != null && Click.ToBool())
                UIElement.SetFocus();
            StaticMethods.SendKeys(UIElement, Text, Click.ToBool());
            //WriteObject(UIElement);
        }
    }

    [Cmdlet(VerbsLifecycle.Invoke, "uiToggle")]
    [OutputType(typeof(AutomationElement))]
    public class InvokeuiToggleCmdlet : Cmdlet
    {
        [Parameter(Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public AutomationElement UIElement { get; set; }
        [Parameter(Position = 1)]
        [Alias("t")]
        public int Timeout { get; set; } = 5000;
        [Parameter()]
        public SwitchParameter On { get; set; }
        [Parameter()]
        public SwitchParameter Off { get; set; }

        protected override void ProcessRecord()
        {
            if (UIElement != null)
            {
                UIElement.SetFocus();
                TogglePattern tp = GetPattern.GetTogglePattern(UIElement);
                if ((!On.ToBool() && !Off.ToBool()) || (On.ToBool() && Off.ToBool()))
                {
                    tp.Toggle();
                }
                else if (Off.ToBool())
                {
                    while (tp.Current.ToggleState != ToggleState.Off && Timeout > 0)
                    {
                        tp.Toggle();
                        Thread.Sleep(1000);
                        Timeout = Timeout - 1000;
                    }
                }
                else if (On.ToBool())
                {
                    while (tp.Current.ToggleState != ToggleState.On && Timeout > 0)
                    {
                        tp.Toggle();
                        Thread.Sleep(1000);
                        Timeout = Timeout - 1000;
                    }
                }
            }
        }
    }

    [Cmdlet(VerbsLifecycle.Invoke, "uiWait")]
    [OutputType(typeof(AutomationElement))]
    public class InvokeuiWaitCmdlet : Cmdlet
    {
        [Parameter(Position = 0, ValueFromPipelineByPropertyName = true, ValueFromPipeline = true)]
        public AutomationElement ParentWin { get; set; }

        [Parameter(Position = 1)]
        public string Name { get; set; }

        [Parameter(Position = 2)]
        public string Type { get; set; }

        [Parameter(Position = 3)]
        public string Class { get; set; }

        [Parameter(Position = 4)]
        public string AutoID { get; set; }

        [Parameter(Mandatory = true)]
        public string For { get; set; }

        [Parameter()]
        public int Timeout { get; set; } = 30000;

        public AutomationElement UIElement { get; set; } = null;

        protected override void ProcessRecord()
        {
            if (For != null && ParentWin != null)
            {
                AutomationElement elem = StaticMethods.FindUnderElement(ParentWin, Name, Class, AutoID, Type, 1000);
                var StartTime = DateTime.Now;
                var EndTime = StartTime.AddMilliseconds(Timeout);
                For = For.ToLower();
                if (For == "appear")
                {
                    while (elem == null && DateTime.Now < EndTime)
                    {
                        try
                        {
                            elem = StaticMethods.FindUnderElement(ParentWin, Name, Class, AutoID, Type, 1000);
                        }
                        catch { }                        
                    }

                    if (elem == null)
                    {
                        WriteError(new ErrorRecord(new TimeoutException("Command timed out before element could appear"), "1", ErrorCategory.InvalidOperation, ParentWin));
                        return;
                    }
                    //else
                    //    WriteObject(elem);
                }
                else if (For == "disappear")
                {
                    StartTime = DateTime.Now;
                    EndTime = StartTime.AddMilliseconds(Timeout);
                    while (elem != null && DateTime.Now < EndTime)
                    {
                        try
                        {
                            elem = StaticMethods.FindUnderElement(ParentWin, Name, Class, AutoID, Type, 1000);
                        }
                        catch { }
                    }
                    if (Timeout <= 0)
                    {
                        WriteError(new ErrorRecord(new TimeoutException("Command timed out before element could disappear"), "1", ErrorCategory.InvalidOperation, ParentWin));
                        return;
                    }
                    //else
                    //    WriteObject(elem);
                }
                else if (For == "visible")
                {
                    elem = StaticMethods.FindUnderElement(ParentWin, Name, Class, AutoID, Type, 1000);
                    while (elem.Current.IsOffscreen && Timeout > 0 && elem != null)
                    {
                        Thread.Sleep(200);
                        Timeout = Timeout - 200;
                    }

                    if (elem == null)
                    {
                        WriteError(new ErrorRecord(new InvalidDataException("The element does not exist yet. Make sure the element exists before checking if its visible ."), "1", ErrorCategory.InvalidOperation, ParentWin));
                        return;
                    }

                    else if (Timeout <= 0 && elem.Current.IsOffscreen)
                    {
                        WriteError(new ErrorRecord(new TimeoutException("Command timed out before element could be visible"), "1", ErrorCategory.InvalidOperation, ParentWin));
                        return;
                    }

                    //else
                    //    WriteObject(elem);
                }
                else if (For == "hidden")
                {
                    elem = StaticMethods.FindUnderElement(ParentWin, Name, Class, AutoID, Type, 1000);
                    while (!elem.Current.IsOffscreen && Timeout > 0 && elem != null)
                    {
                        Thread.Sleep(200);
                        Timeout = Timeout - 200;
                    }

                    if (elem == null)
                    {
                        WriteError(new ErrorRecord(new InvalidDataException("The element does not exist yet. Make sure the element exists before checking if its hidden."), "1", ErrorCategory.InvalidOperation, ParentWin));
                    }

                    else if (Timeout <= 0 && elem.Current.IsOffscreen)
                    {
                        WriteError(new ErrorRecord(new TimeoutException("Command timed out before element could be visible"), "1", ErrorCategory.InvalidOperation, ParentWin));
                    }

                    //else
                    //    WriteObject(elem);
                }
                UIElement = elem;
            }
            else
            {
                WriteError(new ErrorRecord(new InvalidOperationException("Invalid usage of command"), "1", ErrorCategory.InvalidOperation, ParentWin));
                return;
            }
        }
    }

    [Cmdlet(VerbsLifecycle.Invoke, "uiDelay")]
    public class InvokeuiDelayCmdlet : Cmdlet
    {
        [Parameter(Position = 0)]
        public int MilliSecs { get; set; } = 0;
        [Parameter(Position = 1)]
        public int Seconds { get; set; } = 0;

        protected override void ProcessRecord()
        {
            Thread.Sleep(MilliSecs);
            Thread.Sleep(Seconds);
        }
    }

    public static class MouseControl
    {
        [DllImport("user32")]
        static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
        public static void MouseClick(string button)
        {
            switch (button)
            {
                case "left":
                    mouse_event((uint)MouseEventFlags.LEFTDOWN, 0, 0, 0, 0);
                    Thread.Sleep(50);
                    mouse_event((uint)MouseEventFlags.LEFTUP, 0, 0, 0, 0);
                    break;
                case "right":
                    mouse_event((uint)MouseEventFlags.RIGHTDOWN, 0, 0, 0, 0);
                    Thread.Sleep(50);
                    mouse_event((uint)MouseEventFlags.RIGHTUP, 0, 0, 0, 0);
                    break;
                case "middle":
                    mouse_event((uint)MouseEventFlags.MIDDLEDOWN, 0, 0, 0, 0);
                    Thread.Sleep(50);
                    mouse_event((uint)MouseEventFlags.MIDDLEUP, 0, 0, 0, 0);
                    break;
            }
        }
        public static void SetCursorPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }
        public static void DragAndDrop(int StartX, int StartY, int EndX, int EndY)
        {
            SetCursorPos(StartX, StartY);
            mouse_event((uint)MouseEventFlags.LEFTDOWN, 0, 0, 0, 0);
            SetCursorPos(EndX, EndY);
            mouse_event((uint)MouseEventFlags.LEFTUP, 0, 0, 0, 0);
        }
        public static void ButtonDown(string Button)
        {
            switch (Button.ToLower())
            {
                case "left":
                    mouse_event((uint)MouseEventFlags.LEFTDOWN, 0, 0, 0, 0);
                    break;
                case "right":
                    mouse_event((uint)MouseEventFlags.RIGHTDOWN, 0, 0, 0, 0);
                    break;
                case "middle":
                    mouse_event((uint)MouseEventFlags.MIDDLEDOWN, 0, 0, 0, 0);
                    break;
            }
        }
        public static void ButtonUp(string Button)
        {
            switch (Button)
            {
                case "left":
                    mouse_event((uint)MouseEventFlags.LEFTUP, 0, 0, 0, 0);
                    break;
                case "right":
                    mouse_event((uint)MouseEventFlags.RIGHTUP, 0, 0, 0, 0);
                    break;
                case "middle":
                    mouse_event((uint)MouseEventFlags.MIDDLEUP, 0, 0, 0, 0);
                    break;
            }
        }
    }

    public enum MouseEventFlags
    {
        LEFTDOWN = 0x00000002,
        LEFTUP = 0x00000004,
        MIDDLEDOWN = 0x00000020,
        MIDDLEUP = 0x00000040,
        MOVE = 0x00000001,
        ABSOLUTE = 0x00008000,
        RIGHTDOWN = 0x00000008,
        RIGHTUP = 0x00000010
    }

    public class MousePoint
    {
        public MousePoint(int X, int Y)
        {
            Xp = X;
            Yp = Y;
        }
        public int Xp { get; set; }
        public int Yp { get; set; }
    }

    public static class GetPattern
    {
        #region Pattern Definitions
        public static InvokePattern GetInvokePattern(AutomationElement element)
        {
            return element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        }

        public static ValuePattern GetValuePattern(AutomationElement element)
        {
            return element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
        }

        public static TogglePattern GetTogglePattern(AutomationElement element)
        {
            return element.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
        }

        public static ScrollPattern GetScrollPattern(AutomationElement element)
        {
            return element.GetCurrentPattern(ScrollPattern.Pattern) as ScrollPattern;
        }

        public static ScrollItemPattern GetScrollItemPattern(AutomationElement element)
        {
            return element.GetCurrentPattern(ScrollItemPattern.Pattern) as ScrollItemPattern;
        }

        public static TextPattern GetTextPattern(AutomationElement element)
        {
            return element.GetCurrentPattern(TextPattern.Pattern) as TextPattern;
        }

        public static SynchronizedInputPattern GetSynchronizedInputPattern(AutomationElement element)
        {
            return element.GetCurrentPattern(SynchronizedInputPattern.Pattern) as SynchronizedInputPattern;
        }

        public static SelectionPattern GetSelectionPattern(AutomationElement element)
        {
            return element.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;
        }

        public static SelectionItemPattern GetSelectionItemPattern(AutomationElement element)
        {
            return element.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
        }

        public static ExpandCollapsePattern GetExpandCollapsePattern(AutomationElement element)
        {
            return element.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
        }

        public static WindowPattern GetWindowPattern(AutomationElement element)
        {
            return element.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
        }
        #endregion
    }

    public class StaticMethods
    {
        public static AutomationElement FindAtIndex(AutomationElement Elem, string @ElemName, string @ElemClass, string ElemType, string ElemAutoID, int Timeout, int idx = 1)
        {
            ElemName = String.IsNullOrEmpty(ElemName.Trim()) ? "*" : ElemName.Replace("*", ".*").Replace("(", "\\(").Replace(")", "\\)");
            ElemClass = String.IsNullOrEmpty(ElemClass.Trim()) ? "*" : ElemClass.Replace("*", ".*").Replace("(", "\\(").Replace(")", "\\)");
            ElemAutoID = String.IsNullOrEmpty(ElemAutoID.Trim()) ? "*" : ElemAutoID.Replace("*", ".*").Replace("(", "\\(").Replace(")", "\\)");

            Regex NameRegex = new Regex(ElemName, RegexOptions.IgnoreCase);
            Regex ClassRegex = new Regex(ElemClass, RegexOptions.IgnoreCase);
            Regex AutoIDRegex = new Regex(ElemAutoID, RegexOptions.IgnoreCase);

            AutomationElement _ReturnElement = null;

            var StartTime = DateTime.Now;
            var EndTime = StartTime.AddMilliseconds(double.Parse(Timeout.ToString()));
            while (DateTime.Now < EndTime && _ReturnElement == null)
            {
                TreeWalker walker = TreeWalker.ControlViewWalker;
                AutomationElement child = walker.GetFirstChild(Elem);

                int indexCounter = 0;
                while (child != null)
                {
                    //FirstChildren.Add(child);
                    try
                    {
                        if (child.Current.ControlType.ProgrammaticName.ToLower().Split('.')[1] == ElemType.ToLower())
                            if (NameRegex.Match(child.Current.Name).Success)
                                if (ClassRegex.Match(child.Current.ClassName).Success)
                                    if (AutoIDRegex.Match(child.Current.AutomationId).Success)
                                    {
                                        indexCounter += 1;
                                        if (indexCounter == idx)
                                            return child;
                                    }
                        child = walker.GetNextSibling(child);
                    }
                    catch (COMException)
                    {
                        child = null;
                    }
                }
                /*
                List<AutomationElement> TypeFiltered = FirstChildren.Where(x => x.Current.ControlType.ProgrammaticName.ToLower().Split('.')[1] == ElemType.ToLower()).ToList();
                List<AutomationElement> NameFiltered = TypeFiltered.Where(x => NameRegex.Match(x.Current.Name).Success).ToList();
                List<AutomationElement> ClassFiltered = NameFiltered.Where(x => ClassRegex.Match(x.Current.ClassName).Success).ToList();

                if (ClassFiltered.Count >= idx)
                {
                    _ReturnElement = ClassFiltered[idx - 1];
                    return _ReturnElement;
                }
                else
                {
                    _ReturnElement = null;
                }
                }
                */
            }
            return null;
        }

        public static AutomationElement FindUnderElement(AutomationElement ParentWindow, string @ElemName, string @ElemClass, string @ElemAutoID, string ElemType, int Timeout, int idx = 1)
        {
            AutomationElement _ReturnElement = null;

            ElemName = ElemName.Replace("*", ".*").Replace("(", "\\(").Replace(")", "\\)");
            ElemClass = ElemClass.Replace("*", ".*").Replace("(", "\\(").Replace(")", "\\)");
            ElemAutoID = ElemAutoID.Replace("*", ".*").Replace("(", "\\(").Replace(")", "\\)");
            Regex NameRegex = new Regex(ElemName, RegexOptions.IgnoreCase);
            Regex ClassRegex = new Regex(ElemClass, RegexOptions.IgnoreCase);
            Regex AutoIDRegex = new Regex(ElemAutoID, RegexOptions.IgnoreCase);

            var StartTime = DateTime.Now;
            var EndTime = StartTime.AddMilliseconds(double.Parse(Timeout.ToString()));
            while (DateTime.Now <= EndTime && _ReturnElement == null)
            {
                List<AutomationElement> AllDesc = GetAllChildren(ParentWindow); //ParentWindow.FindAll(TreeScope.Descendants, System.Windows.Automation.Condition.TrueCondition).Cast<AutomationElement>().ToList();
                List<AutomationElement> TypeFiltered = AllDesc.Where(x => x.Current.ControlType.ProgrammaticName.ToLower().Split('.')[1] == ElemType.ToLower()).ToList();
                List<AutomationElement> NameFiltered = TypeFiltered.Where(x => NameRegex.Match(x.Current.Name).Success).ToList();
                List<AutomationElement> ClassFiltered = NameFiltered.Where(x => ClassRegex.Match(x.Current.ClassName).Success).ToList();
                List<AutomationElement> AutoIDFiltered = ClassFiltered.Where(x => AutoIDRegex.Match(x.Current.AutomationId).Success).ToList();
                List<AutomationElement> Filtered = AutoIDFiltered;

                _ReturnElement = Filtered.FirstOrDefault();
                //MessageBox.Show(TypeFiltered.Count.ToString() + " : " + NameFiltered.Count.ToString() + " : " + ClassFiltered.Count.ToString() + " : " + AutoIDFiltered.Count.ToString() + " : " + Filtered.Count.ToString());
                //MessageBox.Show(string.Join(",", TypeFiltered.Select(x => x.Current.Name)));

                if (Filtered.Count >= idx)
                {
                    _ReturnElement = Filtered[idx - 1];
                    return _ReturnElement;
                }
                else
                {
                    _ReturnElement = null;
                }
            }

            return _ReturnElement;
        }

        public static List<AutomationElement> GetAllChildren(AutomationElement Parent)
        {
            List<AutomationElement> AllChildren = new List<AutomationElement>();
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement child = walker.GetFirstChild(Parent);

            while (child != null)
            {
                AllChildren.Add(child);
                var RecChildren = GetAllChildren(child);
                child = walker.GetNextSibling(child);
                AllChildren.AddRange(RecChildren);
            }
            return AllChildren;
        }

        public static void SendKeys(AutomationElement element, string Text, bool TryClickElement)
        {
            if (TryClickElement) { ElementClick(element, "left"); }
            System.Windows.Forms.SendKeys.SendWait(Text);
        }

        public static MousePoint ClickablePoint(AutomationElement element)
        {
            if (element == null) throw new ElementNotAvailableException();
            Rect ElemRect = element.Current.BoundingRectangle;
            int X_Center = (int)(ElemRect.X + ElemRect.Width / 2);
            int Y_Center = (int)(ElemRect.Y + ElemRect.Height / 2);
            MousePoint point = new MousePoint(X_Center, Y_Center);
            return point;
        }

        public static void MouseToElem(AutomationElement element)
        {
            MousePoint point = ClickablePoint(element);
            MouseControl.SetCursorPosition(point.Xp, point.Yp);
        }

        public static void ElementClick(AutomationElement element, string MouseButton)
        {
            MouseToElem(element);
            Thread.Sleep(100);
            MouseControl.MouseClick(MouseButton);
        }
    }


    #region Install Tool Specific
    public class Command
    {
        public string Path { get; set; } = ".";
        public string[] Arguments { get; set; }
        public string Action { get; set; }
        public string CtrlType { get; set; }
        public string WinName { get; set; }
        public string ElName { get; set; }
        public string Class { get; set; }
        public string AutoID { get; set; }
        public int Timeout { get; set; }
        public string For { get; set; }
        public string Text { get; set; }
        public bool Click { get; set; }
        public string Button { get; set; }
        public int Seconds { get; set; }
        public int MillSecs { get; set; }
        public bool Focus { get; set; }
        public int DelayBefore { get; set; }
        public int DelayAfter { get; set; }
        public int Index { get; set; }
        public int WinIndex { get; set; }

        public void ElementClick()
        {
            AutomationElement _RootElem = AutomationElement.RootElement;
            AutomationElement ParentWin = StaticMethods.FindAtIndex(_RootElem, WinName, Class, "window", AutoID, Timeout, Index);
            if (ParentWin != null)
            {
                AutomationElement UIElement = StaticMethods.FindUnderElement(ParentWin, ElName, Class, AutoID, CtrlType, Timeout, Index);
                if (UIElement != null)
                {
                    if (Focus)
                        UIElement.SetFocus();
                    StaticMethods.ElementClick(UIElement, Button);
                }                    
            }
        }

        public void SendKeys()
        {
            AutomationElement _RootElem = AutomationElement.RootElement;
            AutomationElement ParentWin = StaticMethods.FindAtIndex(_RootElem, WinName, Class, "window", AutoID, Timeout, Index);
            AutomationElement UIElement = StaticMethods.FindUnderElement(ParentWin, ElName, Class, AutoID, CtrlType, Timeout, Index); ;
            if (UIElement == null && Click)
                throw new ElementNotAvailableException("Element to be clicked before typing isn't available. Unable to execute.");
            
            else if (UIElement != null && Click)
                UIElement.SetFocus();
            StaticMethods.SendKeys(UIElement, Text, Click);
            StaticMethods.SendKeys(UIElement, Text, Click);
        }
    }
    public class BuildTestData
    {
        public string BuildName { get; set; }
        public int Timeout { get; set; }
        public string Type { get; set; }
        public string TCName { get; set; }
        public List<Command> Commands { get; set; }
        public List<CmdResult> CommandResults { get; set; }
        public bool IsComplete { get; set; } = false;
        public TimeSpan TimeTaken
        {
            get
            {
                _TimeTaken = TimeSpan.FromMilliseconds(0);
                foreach (CmdResult cr in CommandResults)
                    _TimeTaken = _TimeTaken + cr.TimeTaken;
                return _TimeTaken;
            }
        }
        private TimeSpan _TimeTaken = TimeSpan.FromSeconds(0);
        public TimeSpan PSTimeTaken { get; set; }
    }
    public class CmdResult
    {
        public string Status { get; set; }
        public string Cmd { get; set; }
        public string Details { get; set; }
        public TimeSpan TimeTaken { get; set; }
    }
    public class BuildData
    {
        public string BuildName { get; set; }
        public int TimeoutMilliseconds { get; set; }
        public string Type { get; set; }
        public string TCName { get; set; }
        public List<Command> Commands { get; set; }
    }
    public class BuildDGData
    {
        public string BuildName { get; set; }
        public string PackType { get; set; }
        public string Date { get; set; }
        public string Size { get; set; }
        public object Tag;
    }
    public class StatusDGData
    {
        public string Build { get; set; }
        public string Status { get; set; }
        public DateTime TestTime { get; set; }
        public TimeSpan TimeTaken { get; set; } = TimeSpan.FromMilliseconds(0);
        public object Tag;
    }
    #endregion
}