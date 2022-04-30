using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Automation;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace AsistenteDeEscritura
{
    public partial class Form2 : Form
    {
        class DynamicBounds
        {
            
            bool m_dirty = true;
            bool m_visible = true;
            int m_left = 0;
            int m_top = 0;
            int m_width = 0;
            int m_height = 0;
            Control m_control;

            public DynamicBounds(Control i_control)
            {
                m_control = i_control;
            }
            public void Apply()
            {
                if (m_visible)
                {
                    if (m_dirty)
                    {
                        m_control.SetBounds(m_left, m_top, m_width, m_height);
                        m_control.Visible = true;
                    }
                }
                else
                {
                    m_control.Visible = false;
                }
                m_dirty = false;
            }

            public bool Update(Word.Window window, Word.Range i_object, int deltaLeft, int deltaTop)
            {
                int rangeLeft = 0;
                int rangeTop = 0;
                int rangeWidth = 0;
                int rangeHeight = 0;
                try
                {
                    window.GetPoint(out rangeLeft, out rangeTop, out rangeWidth, out rangeHeight, i_object);
                    return Update(rangeLeft - deltaLeft, rangeTop - deltaTop, rangeWidth, rangeHeight);
                }
                catch(Exception e)
                {
                    m_dirty = m_visible == true;
                    var te = i_object.Text;
                    m_visible = false;
                    return m_dirty;
                }
            }

            public bool Update(int i_left, int i_top, int i_width, int i_height)
            {
                m_dirty = (i_left != m_left || i_top != m_top || i_width != m_width || i_height != m_height || m_visible == false);
                m_left = i_left;
                m_top = i_top;
                m_width = i_width;
                m_height = i_height;
                m_visible = true;
                return m_dirty;
            }

            public bool UpdateAsInvisible()
            {
                m_dirty = m_visible == true;
                m_visible = false;
                return m_dirty;
            }
        }

        class FlaggedRangeProperties
        {
            public PictureBox rect;
            public DynamicBounds bounds;
            
            public void Invalidate()
            {
                m_isValid = false;
            }

            public void AddReason(IFlagReason i_reason)
            {
                m_isValid = true;
            }

            public bool IsValid()
            {
                return m_isValid;
            }

            private bool m_isValid;

        }

        private List<KeyValuePair<Word.Range, FlaggedRangeProperties>> m_flaggedRanges;
        private Mutex m_flaggedRangeMutex;
        //private Dictionary<Word.Range, FlaggedRangeProperties> m_flaggedRanges;
        
        //private Word.Window m_mainWindow;
        private DynamicBounds m_windowBounds;
        private Word.Window m_lastKnownActiveWindow;
        class RangeComparer : IEqualityComparer<Word.Range>
        {
            public bool Equals(Range x, Range y)
            {
                return x.IsEqual(y);
            }

            public int GetHashCode(Range obj)
            {
                return obj.Start * 1000000 + obj.End + obj.Text.GetHashCode();
                //return obj.GetHashCode();
            }
        }

        public Form2()
        {
            m_flaggedRanges = new List<KeyValuePair<Word.Range, FlaggedRangeProperties>>();
            //m_flaggedRanges = new Dictionary<Word.Range, FlaggedRangeProperties>(new RangeComparer());
            
            InitializeComponent();
            
            m_windowBounds = new DynamicBounds(this);
            m_flaggedRangeMutex = new Mutex();
        }

        public void FlagRange(Word.Range i_range, IFlagReason i_reason)
        {
            if (m_flaggedRanges.Exists(x => x.Key.IsEqual(i_range)))
            //if (m_flaggedRanges.ContainsKey(i_range))
            {
                //do nothing?
                //update properties?
                FlaggedRangeProperties props = m_flaggedRanges.Find(x => x.Key.IsEqual(i_range)).Value;
                props.AddReason(i_reason);
            }
            else
            {
                FlaggedRangeProperties props = new FlaggedRangeProperties();
                props.AddReason(i_reason);
                m_flaggedRangeMutex.WaitOne();
                m_flaggedRanges.Add(new KeyValuePair<Word.Range, FlaggedRangeProperties>(i_range, props));
                m_flaggedRangeMutex.ReleaseMutex();
                //m_flaggedRanges.Add(i_range, props);

            }
        }


        void Application_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            Visible = true;
        }

        void Application_WindowDeactivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            Visible = false;
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        protected override void WndProc(ref Message m)
        {
            const int WM_NCHITTEST = 0x0084;
            const int HTTRANSPARENT = (-1);

            if (m.Msg == WM_NCHITTEST)
            {
                m.Result = (IntPtr)HTTRANSPARENT;
            }
            else
            {
                base.WndProc(ref m);
            }
        }

       

        bool m_processing = false;

        private void RefreshPositionsSync()
        {
            //Visible = false;
            long t1 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            m_windowBounds.Apply();
            m_flaggedRangeMutex.WaitOne();
            foreach (var kv in m_flaggedRanges)
            {
                Word.Range range = kv.Key;
                FlaggedRangeProperties props = kv.Value;
                if (props.bounds != null)
                {
                    props.bounds.Apply();
                }
            }
            m_flaggedRangeMutex.ReleaseMutex();
            long t2 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            long total = t2 - t1;
            if(total > 0)
            {
                int i = 0;
                ++i;
            }
            //Visible = true;
        }

        Word.Window GetMainWindow()
        {
            if(Globals.ThisAddIn.Application.Windows.Count > 0)
            {
                Word.Window activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
                if(activeWindow != m_lastKnownActiveWindow)
                {
                    activeWindow.Application.WindowDeactivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowDeactivateEventHandler(Application_WindowDeactivate);
                    activeWindow.Application.WindowActivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowActivateEventHandler(Application_WindowActivate);
                    m_lastKnownActiveWindow = activeWindow;
                }
                return activeWindow;
            }
            else
            {
                return null;
            }
        }

       // more (rather than less)
            // does not do headers and footers
        private Word.Range GetCurrentlyVisibleRange()
        {
            try
            {
                Word.Window activeWindow = GetMainWindow();
                if(activeWindow == null)
                {
                    return null;
                }
                var left = activeWindow.Application.PointsToPixels(activeWindow.Left);
                var top = activeWindow.Application.PointsToPixels(activeWindow.Top);
                var width = activeWindow.Application.PointsToPixels(activeWindow.Width);
                var height = activeWindow.Application.PointsToPixels(activeWindow.Height);
                var usableWidth = activeWindow.Application.PointsToPixels(activeWindow.UsableWidth);
                var usableHeight = activeWindow.Application.PointsToPixels(activeWindow.UsableHeight);

                var startRangeX = left;// + (width - usableWidth);
                var startRangeY = top;// + (height - usableHeight);

                var endRangeX = startRangeX + width;//usableWidth;
                var endRangeY = startRangeY + height;//usableHeight;

                Word.Range start = (Word.Range)activeWindow.RangeFromPoint((int)startRangeX, (int)startRangeY);
                Word.Range end = (Word.Range)activeWindow.RangeFromPoint((int)endRangeX, (int)endRangeY);
                if(start == null || end == null)
                {
                    return null;
                }
                Word.Range r = activeWindow.Application.ActiveDocument.Range(start.Start, end.Start);

                return r;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private bool RangeIntersects(Word.Range a, Word.Range b)
        {
            if(a == null || b == null)
            {
                return false;
            }
            return a.Start <= b.End && b.Start <= a.End;
        }


        private bool RefreshBoundsAsync()
        {
            bool dirty = false;
            long t1 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            if (m_processing)
            {
                return false;
            }
            m_processing = true;

            
            
            int paneLeft = 0;
            int paneTop = 0;
            int paneWidth = 0;
            int paneHeight = 0;
            Word.Window activewindow = GetMainWindow();
            if(activewindow == null)
            {
                m_processing = false;
                return false;
            }
            activewindow.GetPoint(out paneLeft, out paneTop, out paneWidth, out paneHeight, activewindow.Document);
            var _usableWidth = activewindow.Application.PointsToPixels(activewindow.UsableWidth);
            var _usableHeight = activewindow.Application.PointsToPixels(activewindow.UsableHeight);
            

            //position dialog relative to word insertion point (caret)
            int left = paneLeft;
            int top = paneTop;
            int width = (int)_usableWidth;
            int height = (int)_usableHeight;

            //SetBounds(left, top, width, height);
            bool winDirty = m_windowBounds.Update(left, top, width, height);
            dirty = dirty || winDirty;
            //Location = new System.Drawing.Point(left, top);
            //Size = new System.Drawing.Size(width, height);

            Word.Range visibleRange = GetCurrentlyVisibleRange();

            long t2 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            m_flaggedRangeMutex.WaitOne();
            foreach (var kv in m_flaggedRanges)
            {
                Word.Range range = kv.Key;
                FlaggedRangeProperties props = kv.Value;
                if(props.rect == null)
                {
                    props.rect = new PictureBox();
                    props.rect.BackColor = Color.Red;
                    props.bounds = new DynamicBounds(props.rect);
                    this.Controls.Add(props.rect);
                }
                if (RangeIntersects(visibleRange, range))
                {
                    bool rangeDirty = props.bounds.Update(activewindow, range, left, top);
                    dirty = dirty || rangeDirty;
                }
                else
                {
                    bool rangeDirty = props.bounds.UpdateAsInvisible();
                    dirty = dirty || rangeDirty;
                }
            }
            m_flaggedRangeMutex.ReleaseMutex();
            long t3 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            //this.BringToFront();
            m_processing = false;
            long total1 = t1 - t2;
            long tota2 = t3 - t2;

            return dirty;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            bool dirty = RefreshBoundsAsync();
            if (dirty)
            {
                RefreshPositionsSync();
            }
        }

        public void BeginProcess()
        {
            foreach(var kv in m_flaggedRanges)
            {
                kv.Value.Invalidate();
            }
        }

        public void EndProcess()
        {
            List<Word.Range> toRemove = new List<Range>();
            foreach (var kv in m_flaggedRanges)
            {
                FlaggedRangeProperties props = kv.Value;
                if(!props.IsValid())
                {
                    this.Controls.Remove(props.rect);
                    toRemove.Add(kv.Key);
                }
            }

            foreach(Word.Range range in toRemove)
            {
                m_flaggedRanges.RemoveAll(x => x.Key.IsEqual(range));
            }
        }
    }
}
