using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tulpep.NotificationWindow;

namespace PopupWindow
{
    public class MPopupNotifier : PopupNotifier
    {
        public MPopupNotifier()
        {
            AnimationDuration = 1000;
            AnimationInterval = 1;
            BodyColor = System.Drawing.Color.FromArgb(30, 30, 30);
            BorderColor = System.Drawing.Color.FromArgb(0, 0, 0);
            ContentColor = System.Drawing.Color.FromArgb(255, 255, 255);
            ContentFont = new System.Drawing.Font("ＭＳ ゴシック", 10F);
            ContentHoverColor = System.Drawing.Color.FromArgb(255, 255, 255);
            ContentPadding = Padding.Empty;
            Delay = 15000;
            GradientPower = 20;
            HeaderHeight = 1;
            Scroll = true;
            ShowCloseButton = false;
            ShowGrip = false;
            ShowOptionsButton = true;
        }
    }
}
