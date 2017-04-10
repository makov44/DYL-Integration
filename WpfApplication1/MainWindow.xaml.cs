using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using mshtml;
using System.Threading;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;
using Clipboard = System.Windows.Clipboard;
using MessageBox = System.Windows.MessageBox;
using TextDataFormat = System.Windows.TextDataFormat;
using WebBrowser = System.Windows.Forms.WebBrowser;

namespace WpfApplication1
{       
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string startUrl = @"F:\temp\Roman, Josh - Outlook Web App.htm";
        private string newEmailUrl = @"Untitled Message.htm";
        private HtmlDocument document;
        private object obj = new object();
        private bool newEmailEventHandeled = false;

        public MainWindow()
        {
            InitializeComponent();
            browser.Navigate(startUrl);
            browser.DocumentCompleted += BrowserOnDocumentCompleted;
            
        }
        private void BrowserOnDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs webBrowserDocumentCompletedEventArgs)
        {
            //var serviceProvider = (IServiceProvider)browser.Document;
            //if (serviceProvider != null)
            //{
            //    Guid serviceGuid = new Guid("0002DF05-0000-0000-C000-000000000046");
            //    Guid iid = typeof(SHDocVw.WebBrowser).GUID;
            //    var webBrowserPtr = (SHDocVw.WebBrowser)serviceProvider
            //        .QueryService(ref serviceGuid, ref iid);
            //    if (webBrowserPtr != null)
            //    {
            //        webBrowserPtr.NewWindow3 += WebBrowserPtr_NewWindow3;
            //        webBrowserPtr.NewWindow2 += WebBrowserPtr_NewWindow2; 
            //    }
            //}

            HideScriptErrors(browser, true);

            ((SHDocVw.DWebBrowserEvents2_Event)browser.ActiveXInstance).NewWindow2 += MainWindow_NewWindow2; ;
         
            document = browser.Document;
        }

        private void MainWindow_NewWindow2(ref object ppDisp, ref bool Cancel)
        {
            Cancel = true;
            if(browser.Url.ToString().Contains(startUrl))
                browser.Navigate("https://webmail.allstate.com");
            else if(browser.Url.ToString().Contains("35434534345565654646456"))
                browser.Navigate("35434534345565654646456");
        }

        public void HideScriptErrors(WebBrowser wb, bool Hide)
        {
            FieldInfo fiComWebBrowser = typeof(WebBrowser).GetField("axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fiComWebBrowser == null) return;
            object objComWebBrowser = fiComWebBrowser.GetValue(wb);
            objComWebBrowser?.GetType().InvokeMember("Silent", BindingFlags.SetProperty, null, objComWebBrowser, new object[] { Hide });
        }

        private void New_Click(object sender, RoutedEventArgs e)
        {
            var element = document.GetElementById("newmsgc");
            if (element != null)
            {
                browser.DocumentCompleted += NewEmail_Click;
                element.InvokeMember("click");
            }

        }

        private void NewEmail_Click(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            lock (obj)
            {
                if (!newEmailEventHandeled)
                {
                    newEmailEventHandeled = true;
                    NewEmailHandler();
                    browser.DocumentCompleted -= NewEmail_Click;
                }
            }
        }

        private void NewEmailHandler()
        {
            var divTo = document.GetElementById("divTo");

            if (divTo != null)
            {
                divTo?.Focus();;
            }

            Thread.Sleep(200);
            SendKeys.SendWait("makov.sergey@gmail.com");
            Thread.Sleep(200);

            var txtSubj = document.GetElementById("txtSubj");

            if (txtSubj != null)
            {
                txtSubj?.Focus();
            }

            Thread.Sleep(200);
            SendKeys.SendWait("subject");
            Thread.Sleep(200);

            var win = document?.Window?.Frames?[0];
            if (win == null)
                return;
            
                var innerdoc = win.Document;
                var frame = innerdoc?.GetElementById("ifBdy");
                frame?.Focus();
            

            var body = @"<p><span id=""ms-rterangepaste-start""></span><font face=""Times New Roman"">

</font><h3 style=""background: rgb(244, 244, 244); margin: 1em 0in 7.5pt;""><span style='color: rgb(66, 66, 66); font-size: 21pt; mso-fareast-font-family: ""Times New Roman"";'><font face=""Times New Roman"">Insurance Quote
Requested </font></span></h3><font face=""Times New Roman"">

</font><font face=""Times New Roman"">
 </font><font face=""Times New Roman"">
  </font><font face=""Times New Roman"">
 </font><font face=""Times New Roman"">
</font><table width=""110"" align=""left"" style=""width: 82.5pt; mso-cellspacing: 0in; mso-yfti-tbllook: 1184; mso-table-lspace: 2.25pt; mso-table-rspace: 2.25pt; mso-table-anchor-vertical: paragraph; mso-table-anchor-horizontal: column; mso-table-left: left; mso-padding-alt: 0in 0in 0in 0in;"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr style=""mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;""><td valign=""top"" style=""padding: 3.75pt 7.5pt 7.5pt 0in; border: rgb(0, 0, 0); border-image: none; background-color: transparent;""><font face=""Times New Roman"">
  </font><p style=""margin: 0in 0in 0pt; mso-element: frame; mso-element-frame-hspace: 2.25pt; mso-element-wrap: around; mso-element-anchor-vertical: paragraph; mso-element-anchor-horizontal: column; mso-height-rule: exactly;""><font face=""Times New Roman""><img width=""100"" height=""100"" align=""left"" style=""border: 3px solid white; border-image: none;"" src=""data:image/png;base64,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"" unselectable=""on"" v:shapes=""_x0000_s1026""></font></p><font face=""Times New Roman"">
  </font></td></tr></tbody></table><font face=""Times New Roman"">

</font><p style=""background: rgb(244, 244, 244);""><span style=""font-size: 13pt;""><font face=""Times New Roman"">Your request for
an insurance quote has been received by our agency.<br>
<br>
Your quote is currently being prepared. We will contact you shortly with your
quote.<br>
<br>
To receive your quote now please call us directly at the phone number below.<br>
<br>
We appreciate your business and we look forward to serving your insurance
needs.</font></span></p><font face=""Times New Roman"">

</font><font face=""Times New Roman"">
 </font><font face=""Times New Roman"">
  </font><font face=""Times New Roman"">
  </font><font face=""Times New Roman"">
  </font><font face=""Times New Roman"">
  </font><font face=""Times New Roman"">
 </font><font face=""Times New Roman"">
</font><table width=""600"" style=""width: 6.25in; mso-cellspacing: 0in; mso-yfti-tbllook: 1184; mso-padding-alt: 0in 0in 0in 0in;"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr style=""mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;""><td width=""150"" valign=""top"" style=""padding: 0in; border: rgb(0, 0, 0); border-image: none; width: 112.5pt; background-color: transparent;""><font face=""Times New Roman"">
  </font><p style=""margin: 0in 0in 0pt;""><span style='mso-fareast-font-family: ""Times New Roman"";'><font face=""Times New Roman""><img width=""100"" height=""100"" style=""margin-top: 6px; margin-left: 50px;"" src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAH/1JREFUeNrsfXmQHNd53/f67rl2Z/aY2RvEQUC0QIoCSQAiRPEQRYu2GR+iKKmsEmk5sli29EdY+SPlJFakJKzEFZUiu2KXHFkyHRVdOmhHSmiSEU1REkGKogSRNgCBIEACwrmL3dmdu+983+vumZ6ZntkFCewiqXlTvd1zdPd73++73/d6med5MGhXThMGJBgAMmgDQAaADNoAkAEggzYAZADIoK1zk6JvHv7bh9d+poNYugyUjAKLFy6AaZipTGp4b0pLzUiymHY9zwWMOT36E/wFL9jzQ//I838EwPjXAHDlBKqMOsUix7Rnwd5/E37aes+Y4Dmu6VYbCydK5RcdwTtZyGqg15KgpSRwkzg+p/ten3r/p7oBubjOIunwT6lWpk5cs21q24dnx2dvyySGJgSRJRzX8Vx84V/APfC960D4vu0YkeB7fOEpfL/RQAhMwC3ctzZREP198z3tW5/h9wJyll2v15ZT504/d2zh+N8VK+UnNEgEYHprl5CLaXTZhdICbJ3e+p9v333rv9wxtb3rNyQiNr4sDzcHN9cCGzcLwWju8XPb89+7fO/yvbdhYAAnMBGeCCzhsSRKuJf4XhZk3PAzvsf3osy/w7+o/1n0UtO3XwdvP7l44veeeOWZJ18/9Non1NT4SbouMd0lBYSQXigvwju3Xv/Uh/fdd6coivxzy7XBCTjf8fxjm2823xwvPHYi7/2NQODn8WN3QwEREQQbx8j3KAESEp6OJSSmJUgcJFGQAqBE/r3EARSbEkSfE1izI3Pwidvuv+vxzBMH9h/e/+5RLXcItdoaAfHWBsZSaQm2T1396G+/5yN3cnNCBCT74LnBRby2azFft3IdxzgfuUCdYngOHdMAyJYIgS2hX22shLSrI+qzwMcgBHbFH0/46qVBiOlYAOzdu345V3Mazz538PmtuaGRFeiR1JU6ib1aM20TdCXxvt/Y++sf8qXCatLfB8WLGOx2vcxfwWA4LDRwEmGP6zdgAuMnMY9trA2J2A6RH7dsCouA0bLu4ViD8dOLker1mY5sIknLPTf+yujJhZNfmC8uPJDSU7ES0AaIp3irsk/ZLMO7trz734ylR8FEMHyp8D2koCtt3hOLQsJ8YofSITIv5En/u8DmbfSUgM8w7Yac95lvEcYKkWChzQQu2/Ry0U4w5kLI44ZjgCqqsGf77vsf+/5j/5op3uk44WoDxBTtVQw53lISp68ubN1H78kgh65hkzPA95h8orbc3BAMIZAOj/lAeCFYKB0sUGke22BAgLV5WE2PyoeJH0fcXGhyJERogGAIHsHDmm6+J3qwuXAVjGRy95gN488UWekPiFRW+nbUsi0YErPXZlKZQEW5PmG9VmzhtklJjLqiwZJKihg2FhIBX9z13ViNBcQPgsCatiMEpqnKgEW1VZvdaGoJj8I0D1qD8bhGwTgNhjPZa081TqEaU/sD4or9/X/LsyCh6lly/dA38kU0EuCFwZ8k+naBcUlxm/GGgDqV9qQC3GYM4schfhzp8eO1WF7XJY+spRKgQzW+leb5SpSrKCG0ISC22ZYoSIyFhh6a2sLft9wTGpvjOCBJEmiSOuw6bqwT1QbIUDnVt6NaXYb0eGpU1hUeQ3iRK4YSQ51MSHqgUzm/txHf16/tASEHM3i/FjDoHlWjDg3b4MfN5AH2KTechZSUQoZx3rLiElqyHSor/lm474V72aoi8W3OeF6EPsSYClPRKdJHMDADwVkFEE/oLyE2s0DT9ClVVsG1nTYbQS0pJ+BCdQn+xTcexqBxEZKq3pSadq+rO4XS+rQ/e9Pvl6rL8PF3fwDuue52mK8sNb8bS4/Atx57DJ59+nuQzWVXvdbqxt3vdCtdwlrOSSSVEqrceqMOuZERePCTD8LoUA7qdqOr72SHknpy3JNRdqRVACmmyn07WBJWYDozN6mCCgY0Ag+jZdQoQKqadfhfL/8DnFueh3Qi3eXjdxLJP33thKNBlZbOwl1v34eSkIYlodT8Li2n4cixV+Hpl56FmZmZt5gW83qc7nWFEOEvy7UK5FJZuP9jHwM5m4c60QiiRt9/l9RTY8XKim7YVr1TztoASdmJ/hJiGjCkpCdkDHSMzrglOKYUyWg6C7Iscwm51I3SDqSOEorfV0pfNPsHFmzbuR3uyfwW5EfG1t19NhwLhDraSLShLvev2ukTGn1d1sbed/2dI4Xh/KlOXmwDZBTG+ruCrggZbWgyqmrCm3rNwNECE42/KArALkdyn4yl0ApwpAggMr5QUQO6lJR9XndAaOySCU3XHiKQeBEnHz0tOaFpI6oqn+orIWe98z1vRoiXhFJS0zSOmtBB7fCyJnKJjRwsSAL9qCe4xCmGZUClUeVGXlM0lKgEN4S91UVwLvOaLBAFREIN3TAaUCwvg67rlx+QjsuTkyEZQlOlCYx1qWt6ZVCV/5+fPZ23bBtdX7//u7bsisll6U7foFAGYSKlJXNNQGLGS4BQOkVDw886AAnzRCu1EpQRiMlsAd4+9zagAOns0nl4ff4ET82MonEm17mX18Vd+2CwkiC26fL8SB62TW+BsewIrHfAT2OXLIGPh+fmOuf/AnKQup0cnZyo2BVQJbW3hNhvWH3sB0blOoynU2l+2TDaZh3MYhEgng1JKdFhY/xz5ksXID80Br/z3o/Ars3XwfTIJOeSxXIRjp1/A558+Rn43qH9kNaS6B5qPAXR1cSWDFFWNcqh99xyN27vf8se1moJSIhMXsW4KhycqBbxImkVcnaUulxoHK+BmOyjspjYm6U8ZqNaSUz4STGKZOMNBEWjNnN8GxK5F3XuQnkRdsxsg3/76w/BlvFNsFxf4Z4JxSIkUTdffRPcsmMPfPXZv4GvfP9RkDGIClP77f1sUUDskBBZUro5c52bjQxJaph1qWzfsGsYoSf0ZN5F15cpYm9AnHxvrjJMDLqGUpMpMekTuMegiTMAgWUia5OQutWAoVQGPv2+34XJoXE4cuZo2/cGenDLlRUuFR+9+V44vXwO/v4fn0aRHO2yKHTtMN9FEuLFeGHr0bxmpop12TmRibHnhGmX9FiqwNDpFRN9AMmpuZ43VzwFhtXsVJgQ65WqbzjczegCpFqvwR079sFsZgKOnz/RO9axyqCLGty85Qb4wbEXwCFpE9o7LUisCZLAAbly5uHX2peUlMw7KzRl3cft/dUb3993HgSj7ynK8MqS3OVltXxxjFCEQEIinENeV04fhkajARa6pP1UfLG0zEHJpoag1Ki0eVK+hAitgTPok8TYmEzxWloykSwYwzVwhpK9AXns+W/3JlKlCHt37J4gMPo1LiHIwUzqAEQWOFiG2QDT6g+IwmSoGVXuPot4HoHbKSGU3o5rR149AkePHAU9oV9Wwx5rOywLEskE7N27lycR+0qIlhpVPEmVXGZE+9l2Vhm5sacqwe90Vc+v1qkGEp0IKHRIiCiLcL66AJU6Etq2uCHv2QQdzpfxt04VkugidhQPcGlzelSmPPLNr8Ef/7fPw5arrlr3iqKVcgkmxgrw1Ncfh5HcSH8JUZJZtZosuLZ4IuogtQGi9EFVlaTEkI4GYDVAXCMw6u0qLaHq8IvaWThTOu+rLnRRWY9Yy0ROO7ZyEiz07ERR6KIr4xIST+0tb98Kv/F7H4Cp/NS6R+rkuCi2tKap8KSWlFE+xhzRPRGlVRsCgtgznYYiKIyhyzuyFgmBQELaYk509ZaMZThUfBVu1fZivGLGqhRNVOBcdR5+vnKM58J8W9QRhkg0kRXvSUmyDIl0iqus9QZEcAQeqa9FMmmiSpTYuOU1cHxKj/S7E48s1Uppoj5B6Q0e/Ll2rDGjqLnuNHx1JXZfi1zanxUPw7XZ7by2qU7gddoPKQk/Kr4MC9YiZNWh+IEjIHYACM2j2IH6o0SjaZrQqDfQeTDWP7lIqRNT5HFILzqF/aQYK6Em8+VGGbWAFA8I9Cjgcm0bVFmbUBTVVymcuzvnDggQHWpugxvdOECSYgIWzEX4weJLcNv4HnDAa0sIKYIM540L8HLpMAePgI3rEUlIOGgD++JEAKFMgY39pQmi9QbEdhyaJeN9gh504rSigBHVFNqRiVK1BMztYdRtOd6DMlBBpDR9EtUW1Jw6vyGLpNw5kQI3GHkTBAWPJS/WKUwgaD8tHYTZ9CTMJiahitdr2hlJh9cqR2HRKUJGTfX06RlGuGRfwsxACIhnV+GOvbfCje/Y1cwnXU731gvnRryQn10+H68mNCiZZV4AGB/JOzzTocpqgRKMjuz2SJ30qi5wPNBkbZJsg8nQbUWITMtGTnSaM2ZUd2RjrHKmdg6KZhF/L7V75UHqh6oViXNX7GUoOSmo2bVIBxyUoPOwbC2h9Ji9XXB7CZZQ0kwEoIqBpBMMnCaMx8ayMC0V/M8uxuv1OhK4kVq/sGihLSMQvJFxnCr4Y7U9t1nvXDIqPY07UUXXdNAUdUIFEXTWw4aIUI29gOhWIZWczF9A8Vp8owyuyXg9qxgUM3iePxlDx+/T7obrZ94FKqofN6zV8ppVMpw7KJdTYGPg1TyI5gYkS4R9SMwdI7ubEhfrOKRMmLLy8Pwrp7mebpeEyptOGDanZlkwc8DnNaBZM9acrmV+ap3uSy48TUwpSQYTkxlQ0CkxGnbfQgueaERtIyly3sEg2mRWD6Ne7Z6gonKY+vISzMvObGnEhjyGIqP5IVBVjNYRENtXm7hH44oI7NqMBpuR/iRngKpDgr3DBc0vj8FdzfZNVrTj9J60nS72d1TIPFXw/KVGPcgYvPUAkPpB1yW7JfoqHjc8FinFHxzzjd4jQ4o+aGSvlpbLcGZxCQ6+Mg/DEwLkx4bBqPeXUGaTJyqO0pgNy+6hsooz7WBgL1eqZcgoYx/75Z03vWfnjum2Kj0iNqWZbKpjckRuNiqU3AvAcGhpgesDQaLMq0zi3YmIvfK3VdUL3jeVVC6dTQikgs/2CQEgzAdACACRmscBeLgnxkwmczAzlYP5+QnYf/Aw0uw0TE0hKDW3JyieyFNABdQEGbQ7pXiyiGab11RvmCDI9l33f/jOr85O5JpcTOrHL+HxpYP2NG/R3LvQVGWhqvLcyPsrtPld9VUv9dfl1ZRIC9f/3I3aWiEYoxBWoQCMjyfhV4ZugG//8EU4yRZgfHQITMONxUS0JGRmIesxJ2t5Vg9AoiU5eFyp1YZ+8+6bv0lgOG7r60670K5f/N6xwK74JaQEsNeVArmSWpydYE3bEpmM8lhztVdYYA6e754TYLIK8N5dN8B3XnoGyqkSag0t1tsTMV6jCF2R1Dy66CfiVZbUSnNXaw3YNFv49O7rt6U8z1c/nSGL58Xr+rCiNRyHEFFzV+qixpD4AjFOBJSOKT+/0I/mO0Je9NpT7iaag6FhAa4evxoOnX4JJmclsM0YKtmMq10EpNBwemR7zaLe7F5p2YRb98x+lHRlw/LapmFaHQndQ6+d0wJjHa73cLl0tJbBXclSIrCotLSD0izjCxwTgbVi6Tba4LZpYgpeLR7GuK2GKk+KCbZdzp0ofXm/6DBOQlQ/JiA7oKasmfxYdhsExjtSU93Ubr6Yeh2S0gFO0PFQpNmVDEgTFK8jhmhVVPoaKiimdlv2w2ubO0IpSQuQUtJQKp+kiLzLdlI85pLHJooFqdd8CNPsIKflgC6zCU2VWyT2YuKoZpml1+QQf/2Hr5rcYFUUGT7Pe3PO6Xo6AQzaJUMICqjDIupWpwI7EsUq0k+ytxqSjmKSWsPwswYdswUsWAHAZDHPFJE8pqUuQPTA4bQ9mhdXhhVVbSsW7iRSSmPQLrCXvlVRmpdWHO5ytnJGAMNDIqSljZUow6GpAq/bLiKBKJ6hhaENw/SDVzcmGYtRvsO8iVK9UogFRJX8iFFwDZB1bULTNB5jhOX1rJncY1z/fOvJV+D0uWVI6EqkeLpj35GGuJhWqTbgnTs3we53zEHVaI1oSGVw5PAxeOXln4KuJy4pQ7BImifUWs3VUsHnLlW2Y8T4rpv3QXYozQPibolmqGVUsKsuX8YBLuvSMA3DQNWWmdw6ehVN/B3qAqRoZ3yuxB9ODWenEqrE7Uczt8NC4BhezIM/+eozsP8nr0F+NNNzYG86JsD7nTq9CP/q0/fAHbs3QSOS2sroAuzf/0P495/9LMzObbrkdslb5QPDNDkoj3z1K1AY3QGWEc9uCUUHysDzuR+n2780BRsSgpwfEYbH4uMQ26e+g6NPqcqUipxo2y3KhuMWwC8Lyg4nYfPcOHJJ8rLoc9KxmbTWzBpE29zcFvjIx/8AJicmeQ3U+gaQ5IUuBtLp25Tm+sioCZAxBkHJsPAV52nRFEXdaVQOnvt59Ra4pRuQmZx/uWXJgVxGnSLNZFKAF8OB/EEACKAoSUi3yxRdCK062Y4ZYX8lkpYAVdM2ABABVFVrFsIx1p7Rbs7/0IwnulImzRHFdHF4KAv/9KOfHHru298//OCvfqIbkJVA9JYb6PbqWoHr0Vj3iPGkmoGACDTbtQogdAlK11NwKaN1liVxbQNnQrO+t/MWLuqCeq3Gy4o2QkLovqGlZjHz+/SbBDpFMpOhbpVA9YR2FShjlI5e7HcfeeLll3/4s9djVVa54RvniqFKuqaPhwYtrpF0mGjxRclfMhw3g0hrABeXawiGA+MjKVSDMiyX6zC/VIZUQoVsRoO+z5thrVWsIuteYkkZZj935q0zIOCvb4x4TF2BH9kQVeFLoZfR5oiiHGEmF/JzBfjHp38CCMb+KAXaFZuxyFEXHWskk0oUAHquKEAJQUAcj6ss6KwsxJOqdRNKlQbsvnYOjfJmuGZrATJJFc4ulOHg0XPw5P6jcPj1BRjLJrnUxMUbHhObCT0hEgtYeN+bdl0H27dtBUVRYCOeIOSghhgbyfK+xNGIhAd5DjQBgwnLBJVpTTDS+SGwig34m8898mP86MfQKzCkpKPrWqjvhNFUUtehT2RtYTBAEiJI7SqLJKPWsHhHH7r/Fvjw3ddBGoEoVQx+zszEMNx202b4rbt2whe/9jw89vQhBCXBQewq90F12FxrIUTVFUB+JMO3jWw0rWDb8TaWa3oERJc0MAwLbNnhOfvczChINQ++9Mkv1t44cuLLNDXUExBp6CpOtCSDyVRCgX4SYllU0ozBDamsSIU6zxI3bPj9+26CT37wJnjjzDKcml/xM72RZ2LlhhLwRw/eDlX87VMvHIN8rttTEyQRQsUg8mi/lTiynSsjjy8KvfNi1M+Uio6HkoTc5AiItgC/eP549Zuf+1r16D8dfQZ/8Wjnee1GvbgCtboBhfHsZCal9wXEtG1OLEZV3hEbUkFVtf2qMXjvns1w9OQFqNTMWDE7PV/iBP61W7fDCwfP8BlGWepYlRVZViywK+nRZmu1NAxSUsK78OLZxpnvHim/tv/V8ouP7xcc8CgI/Axupb6APHDfe8FBQ3lmfmnSDjKQPavcTSopELh0RCvviLCzk1n+HJOllTr0s7fnLpQgk5BhCnXqiXMl9DrELpXlelHXkrVz4JVAdq9/cCmZovPlB//rMTK7RB7cXsDtj3E7FXdeGyCvHjvNk2FoPyYTmtJ34AaqLHqaD+M2JFrlLvPOkFE3TLu/B4SWr27YXBKY2O0c0LW9HiBQQVy1WulZhX95ed/j902l0jTr1zuwxVYo5InGP8PtQGDAf9Dv2m2A/PCnh9FNrcDte3ZOyHL/zJ0PCPMlJAKIrNDyNBNW0MOyUNpI4nreHOl/AaVouWZhoKV01QPTtd0eHPG3//M78JnPPQwz01NrqqW9pElFw+TM8d///E9hy5bNfX87M1OA99x17+efffIbB9Zy7fZHa6QS3KiPZNOTq3bKIpUldgGio693pmjAqYUKTOQSULWcnjKd1CR441wFijUHhpJKlzSShDg9EmKymoTNb7sR5uZm1r1CkZisWllZ05MhzXoD7vv476NFd+DZxx+7OEAYk2jyAt1UffUqdwTEoQewkIcVoaSCErJUNuDQL0owOaLznFccB9NHNVRXB/F3ILR7ak31F7EhXWmHbA6uvf4GmJ6a4A/UXF+7IcDK8oUgBurfFJmBuVLbdO8HHoLNUzvhwAvPQrE4DydOHV4dEFqdJ8vCSDKhjfYNoAkQm6afSDqkdh7GNxqqnwPHS3Dt5mFUSwzBczpneCCjS/Da2TK8eqaODKA1nybX7mVJEK3Z8CL3pyfrNOo1vq176gSZxKCUTcyT8zrppKlUm+gVTr3+Bly36xbYs+/9UC4vw/z8ydUBcakCT2KFREJLhQFYL9+7YXpcQuI4O5mQ4OyKCS8eXYFbd+agZnbWJzGgR5+9dKwMdfQ9crocq3YoDnEiOaCwP3T/Wq0OF5aKaFhTPI2x3oAsF5fRsbBhNTpRS+f0vLCyBHUoYr+XQFJlmN62aXVALEukiZdCEKRDnD1mfrEE1FFlCVK8qqFGRWzPHy1BYVSHubEEqqeWWkmoApxcNODo+QZKh8LD8DjFxCKBYbQ/JEzbr94G9ykJSGfSABuQyzIadcii2vS8eDqFIkJFIrqmThimBbrjPyTAMU0+p7IqIA4in1TkvKqqzdRAF9eGRt2mZWpyl6vasiUizKOUnF22YcsUqh6rdTERPbiFMqobhx4zIfWUeUGUm0ad+hL2h1Ze79g6w7eNbDSkhgU9J+JcTiNyXrSCjFqCShZWY532FVTMgHQqMS3KClApUVx6Qgw8KpIQmpMIJYSn2CnhaPkZWJrWnM0n4eqZDNTowaUR4CyPJp5U/kTSKt4oqUqx8Q5XWaTe3LB+uNWfsrHxa2+9fpFh0DQ06pqmjdOckceEVQPadpWFOlFTtSlaEENTprYT3wk+H2z7mV6arKF8FLVcRoGxIRWGUxJoKCFzeR1VkszVVVS1UeHCVD4Ft14/Dj8+ssy/dxyaK8HOq6JfWU4TY/xRgAIvq/F4ljcuFr6yG2kzRVezjiYMNUR3hQnsItxedFn1hD5FmdVqvfeYSUwpRULcW0VR3DSVhh1zaXRzNUjpkp+TInHGH5FBZx2LSamTBqrH63eMwKbpDCxg3HK+2MDYpQHnlxr8xnQdWrlL9yD1yCvNnXZjJjDYmFrh8Jlka7BdpGlURcspjI0JjrsieBcDCOoFXdPyDk+vxxdSUaaCBKJUdzgH337jBFy3dRidAcY5vY76pWY3U4LAZCF2PDSWCgKbSqkYiOqwgxETOAhKDQ6iy/z6mSqqKl/MCXzyWCy3ZdzT6AcOKQDOxuEBS1VkFsPrO2FKzKSqmoAOyigeviawiwAE3UeGKivfVkgdM/FCDhPp/vfsmoA978jDEhpvqrzgM4cXOb9O1T1GsP6AbNKOzTnYtikLPz1chL9/4TQvYw0XQ4W1TWNpBo8/9QN49Gtf4Z7OeqdOLMtfkPPJP3gIdmyfhlK5t6TQ0g8KIBVZzlMWRFyFPu3zIa6T1XRt0uVeTY8FoEiacs2G3LAKb7s6B8WaBxYFiPJbHygJ1nLdo1gI9t0wASuI/OJ8w5+VQ/0U9klHQ3kcA60nvvsszM5ete6ZX8uyeED6wY88AKo83ZNWfEw285dqq9pE2agg012EhEggjKmangzXfMQCwl09h9sFsjk2qSXp0haqkYdeRgncelUOjNoFqKPOomXEYZ9o0drY5DTc9WsfgsIEPSBgfRUXV7crRVD1JNSc/mEQOSKSrFJBQ4EWgcoX8/9DxFRykmIQZIDewQ5fguXxLCyl2mnN+KUuMuDqidxEynHhntxdx231ia/ewjeW2eCb57nrDAi69ZbZXJzUJ6ENtHqbvEdVVSY9ZDVBkNcOiKxIE4rsA9JrjKQeqg0HltAAV5BQhutelqoPFW9UxPsslC0//vDa+1SvG1AsFiGRTG8IIOXlZV7aBNCbVhCofkowqqo+zsvdPXHtgCREZVLSdO73Oz1dXgzkEjLsQcObQIssXiYFLjg2TGNEn9+Z5/P2BhqYsE919IyzuRG45peugZGR8UBm1ldl1Stl0PQEj9f6Te+Ha2IURcuXawZqlouQEEXXpyT0L/vZkJpBmVoZ7t1V4C5fmx94SRN4/mpXRQVYqbucEULsFxddeOcNu+GGm3ZvaHxIhX8rK+6qA3FcRmWn+Wu2zbJ0Ju2tHZCEPkNemWn1N1RUj2XVLz8lDOxENZL/aaus92BjUycsXOy6+vw+GfZEIjGx+5pNw7nscHHtKiuR5ifzBZ4b+4/SVneRMQC17SujL6tlC/zVZwIcOHjUzQ5lYpl9dtNkNyAvv/Tca3v37QT06NY7o/3/daOH8C2vFF//1rf/bkXT9Fgt+5v33OaDF50Y2rRlx4eeeunnj9IZlbINcAUvY/5/pdFDZqZnGPzhQ5/9+pc+/0f39ZayoKg8+uGJ40cO/O/HHvfGcqQbheZa9MH25jZ6okV2hMHxYwZ886///EdrMk1RCaGc0NTslu9+57mf3zE6JsHZUw6wwX/LfdN2JZkUYTIP8M8/+qn6t/7Hn+4lq3BREkLt9Mljf/3pBz5oYNwF05tEXirqDrh9zRv3vDDwGM6JUEAw/tN//AtAML6BpH3lTUkINiop/8sb3nXXHf/hT76UvvadsyJlYysl4F6N5w0sS0w2yV+PL5KnCpBKAyzMA3zx4X9X+osvfOYMfvUA+CWksJqExAFC7R24fSGRGt70wfsflG+985/l5rb8kpZKp3h23Rt4YF1uLc/xGTaq+detAy9+b/nrf/VntVcPHaClBv8Fty+vruL6AxKC8oe4XUNvClNzLJUZFvi/SR0g0gGIj4hhNLzTJ487tmWSKVgIgHhkbTZndUCo0T+Ruhe3d+NGa6nVwO4MEIlJWYFf4b4c2AuyG0fW7gSsDZCwUaHWFtyGAZoPwhm0bkDoGYm/wG3+4r2yGEAG7QoIJAckGAAyaANABoAM2gCQASCDNgBkAMigrXP7vwIMAHUtkreg2s4jAAAAAElFTkSuQmCC"" unselectable=""on"" v:shapes=""_x0000_i1025""></font></span></p><font face=""Times New Roman"">
  </font></td><td width=""12"" valign=""top"" style=""padding: 0in; border: rgb(0, 0, 0); border-image: none; width: 9pt; background-color: transparent;""><font face=""Times New Roman"">
  </font><p style=""margin: 0in 0in 0pt;""><span style='mso-fareast-font-family: ""Times New Roman"";'><font face=""Times New Roman"">&nbsp;</font></span></p><font face=""Times New Roman"">
  </font></td><td width=""200"" valign=""top"" style=""padding: 0in; border: rgb(0, 0, 0); border-image: none; width: 150pt; background-color: transparent;""><font face=""Times New Roman"">
  </font><p style=""margin: 0in 0in 0pt;""><em><b><span style='color: rgb(66, 66, 66); font-size: 9.5pt; mso-fareast-font-family: ""Times New Roman"";'><font face=""Times New Roman"">Email: </font></span></b></em></p><font face=""Times New Roman"">
  </font><p style=""margin: 0in 0in 0pt;""><font face=""Times New Roman""><em><b><span style='color: rgb(66, 66, 66); font-size: 9.5pt; mso-fareast-font-family: ""Times New Roman"";'>Phone: (781) 864-5992</span></b></em><span style='color: rgb(66, 66, 66); font-size: 9.5pt; mso-fareast-font-family: ""Times New Roman"";'>
  </span></font></p><font face=""Times New Roman"">
  </font><p style=""margin: 0in 0in 0pt; line-height: 10.8pt;""><span style='color: rgb(134, 134, 134); font-size: 9.5pt; mso-fareast-font-family: ""Times New Roman"";'><br><font face=""Times New Roman"">
  Cambridge, MA 02141 </font></span></p><font face=""Times New Roman"">
  </font></td><td valign=""top"" style=""padding: 0in; border: rgb(0, 0, 0); border-image: none; background-color: transparent;""><font face=""Times New Roman"">
  </font><font face=""Times New Roman"">
   </font><font face=""Times New Roman"">
    </font><font face=""Times New Roman"">
   </font><font face=""Times New Roman"">
   </font><font face=""Times New Roman"">
    </font><font face=""Times New Roman"">
   </font><font face=""Times New Roman"">
  </font><table style=""float: right; mso-cellspacing: 0in; mso-yfti-tbllook: 1184; mso-padding-alt: 0in 0in 0in 0in;"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr style=""mso-yfti-irow: 0; mso-yfti-firstrow: yes;""><td style=""padding: 0in; border: rgb(0, 0, 0); border-image: none; background-color: transparent;""><font face=""Times New Roman"">
    </font><p style=""margin: 11.25pt 0in 0pt;""><span style='mso-fareast-font-family: ""Times New Roman"";'><a href=""http://dial-mail.com/email/8405277306/callnow""><span style=""color: blue; text-decoration: none; text-underline: none;""><span style=""mso-ignore: vglayout;""><font face=""Times New Roman""><img width=""138"" height=""43"" src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAIoAAAArCAMAAACQGeAdAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAA2ZpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMC1jMDYwIDYxLjEzNDc3NywgMjAxMC8wMi8xMi0xNzozMjowMCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDpGQjdGMTE3NDA3MjA2ODExQjJCNzg2MEYxNEZDREU1NCIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDowRTIxRDdEQjMxMDMxMUUyQkNGM0VDNzYxOUZGMDQ4MSIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDowRTIxRDdEQTMxMDMxMUUyQkNGM0VDNzYxOUZGMDQ4MSIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ1M1IE1hY2ludG9zaCI+IDx4bXBNTTpEZXJpdmVkRnJvbSBzdFJlZjppbnN0YW5jZUlEPSJ4bXAuaWlkOjAyODAxMTc0MDcyMDY4MTE4NjIwOTE1RjI2OTFBQThDIiBzdFJlZjpkb2N1bWVudElEPSJ4bXAuZGlkOkZCN0YxMTc0MDcyMDY4MTFCMkI3ODYwRjE0RkNERTU0Ii8+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+RFhK6wAAAwBQTFRFOZMNOnYi3+XdDjoAz9nMFlMAq9GZOIkQj6aIr72qJmQA4/Hdgb1mv8q7NoMOcqtVMq8Af7FmnMaINXERpsSZNooMfpZ3ImkAqs2Zjr13jsJ3lK6BM3MN8PbuJ1wAKWcR7/Lux+O7XZREV6YzGlgAfq1mLGQFGVIAJHIAgJp3W4M+1enML1ciudqqPbINSJoiY6NEbolmgLlmKV8CNW0QT3JE8PTuLWkFZK1Ei7V34+/dLGEGWoxEHmEAHV0ALGoRuNSqNX0NIm4AgLZmc5NmgqJ3n7GZT3ZETm5EK4cAHmUAFUsAOZkNYp1EK4QAJ34AFk0AJHoAobiRLW0FQbYRf5l3qcqZOpQROqENOp4ND0AAPa4NVJQzn7KZZrJER5YiOIsNjbt3EUUAcrNVjbl3DTYAQK4RF1kAPJwREkEA4u3dOqYMOI8OIWUAQGkzOrYJPaoOobaZgsBmJ14RO6oMPKUOKWEEXntVxty7ED4AKWYExNa7q9SZPrYONHgOQrISM3kMNrMEE0kAG2IAKWMDLWYIE04ANGsPRo0iJoAAJmEB////MKUAL6EAJ2oAMKMAKHAAMKcAL58AJ2wALpsAKHMAL50AJmgAJmYAJ24AMakAJWIAMasAJWAAMawAMa4AJV8ALpgAM2oOJV0AQL4OLpYALZQALZEAKHUAKXcALI8AKnwAKXoALIwAKoIALIkAKn8A9/f3v8y7v8273+bdb4xm8fjuFEcAbodmF08A4OfdkMd3r8CqJl4BpryWLGYEf7RmM7ABVZkzlrmIJ2UCQb4Qv8678vPx8/Ty6e3mLnEFMq0BOLQHn7SZKX8AaZdVGloA2+TVU301LncRX4BVoLeQPLkK1t7PttGqNa8EQb8PLmkJJW8Ai7h3dLlVSZ8ihql3M60C7fLqTXktgJ5pWIE6Uo8zH1ARG1wAH0wRpLqVmbKHPmAzlbOIn6+Zn7WOP6gQ4evdMmgMMmoMK2gDWIVEgqFsJHgAKooAMqoBO60MVYVEkat9eqtmJG0ANbIES+mJdwAAALJ0Uk5T////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////AAsuQ+IAAAUJSURBVHjazJhpWFRlFMevDCSEgDg4kyzTkJOA4taAaSwFqcGgbGEqhgl6RQwwozIyl1BMe9vLVq2wgGGGAcaFRVHTbLey3cqyzbKy3fZS/+d9Z7nUF770zP194L7vOeee87vvhefhuZLT6Xz++5EpBl/yRtSL0JCczq/N8xY++5QPeX+htDaKVApPrnrnln2Dfci+00uKXvrbKb1gXnVw8OO+5nTR3nelt4cuefkxn7Pp4KLPpWM/vLdJBXxSZJBSiv54RA2s2CsZVtyjCsZcLhnG3K4KuMpdqmAOVObcoQr+JJV7VcF1ULnmAVXAVe5XBVdD5fqGfhDM2NyGaxmLbfi/4CqNfVm0C3PZqF19gghdtgEqz5U38P3NKPmIFnOxiG38DyhlwSId3NgvZl0oGWY19WF4JxNEKKNulW3lYk8q2wqwiGXeoAJSYVOamqCyfV1TfyCV2c1KjsLEv7qsLJC9skYRJpWCQTRV7EmF/dXcPJwxb1DBIIob32yeSCrN/WE2qdiUnIUOXwZssNlSA9dMHE0NR89H2K2ytdxTxlhvgbjyYCSeITjS3YarsKW/cxWbjXeKRaNI9LPZjuIXwGabyVind/A0qExzKEHDhLJsWv1miXS9qn8cDqgkF1yAqQGijIwZGzCWXxAcGyxKn3G1oVJsB5BKDu04kxw4ReMhBzWucczHw5R5BnOVdgXozd6yuDaREYnFxb14YZntF0HFylVECiozGCspZcZAHsR+R3HidgwXeZSWTGEsZBxU6mZ2UjaZ1NpjGKtupzPakTsEPUyeyRJUpGwFV6EoOdO9sxzPo4jfumxSiecqInMOei1nbAILTKBgJXyL9dmlaJ7H85dC5VyMhY5f3TCRPYwDz6P5BZ04y179KMa+sngmc5VcBVwl3r0LG8b/nPy6crkK+m/tERlSSaVc4sd49oBxzM2MNJ4nlUPh/PX5dUGlJCc3ly6WMMYCX6OzNNZAsPZVz+RKqFS2KcGT6DJd60tc7f3K2qCSdeRpTO0RKaiMtyzGXvMhBb0qujSeR+lUU9uP4u5Svmnjl4EQ+JYFVeOEIFXnHVx5EiodSjDDeDEt9g/ZjzbLNInUrINUTuEpQ/I9ZeMt2Gfll1KQBmgE6TwfTmM7wjwq/qaODjyZztRxI2M7mS4uhg4s67h38FD8FzfPrmT9BNw33W5fwCJ+QZuKqnD6aScVPankiLIroJJmT9No9At4EAOWIhx2OFzkSSXObl9OKhWhOOrJ9un0OtPRGJxIpxM3avTewaQysFWJ/Tyj66in1sS4X1BF62KPiijjKnzFVVrX73TVThZRrtLaeqW4+3z329PE85BRc4qMgmoVg0OhEtrSF62OZEJ0ctwX/rirm5q1kEoVOoZ0iSKodKfz1SQRPJLgRzN0J0QUpd/F4RrB725JDaKWy2QTQjHk0EJG3fmKsVzF+i/iTRpZlrtMVqu+QpY16bJca7XWynKVFcsuV1GPLKeLlckVNKFE1pj0IorSHlcd7rZW5VDLNJ5C9xz0lj0diI2hI6CyURVk4FQy7lYFGTiVjIdVwWqorL5PFZDKyjtVwcoUqDyoCqJxKtEPqQKopGgPnK0Gos3SMe1Nj6qAA1qDVPjzB7dt9j31Sb9KToN2xOYnfM1arfl1ybnHrK2/YfeTPmS3uT4pJYq+2+4x/JSUtMWH3Jq0xfwN/4TsdBaO9OnHbMNnUZ/C4owAAwAHeuEsqIK1EgAAAABJRU5ErkJggg=="" border=""0"" unselectable=""on"" v:shapes=""_x0000_i1026""></font></span></span></a></span></p><font face=""Times New Roman"">
    </font></td></tr><tr style=""mso-yfti-irow: 1; mso-yfti-lastrow: yes;""><td style=""padding: 9.75pt 0in 0in 3.75pt; border: rgb(0, 0, 0); border-image: none; background-color: transparent;""><font face=""Times New Roman"">
    </font><p style=""margin: 11.25pt 0in 0pt;""><span style='color: rgb(151, 151, 151); mso-fareast-font-family: ""Times New Roman"";'><font face=""Times New Roman"">Number On File</font></span><span style='mso-fareast-font-family: ""Times New Roman"";'><br><font face=""Times New Roman"">
    </font></span><font face=""Times New Roman""><b><span style='color: black; font-size: 15.5pt; mso-fareast-font-family: ""Times New Roman"";'>(614) 333-2217</span></b><span style='mso-fareast-font-family: ""Times New Roman"";'> </span></font></p><font face=""Times New Roman"">
    </font></td></tr></tbody></table><font face=""Times New Roman"">
  </font></td></tr></tbody></table><font face=""Times New Roman"">

</font><span id=""ms-rterangepaste-end""></span></p>";

            body = body.Replace("\"\"", "\"");
            Clipboard.SetText(body, TextDataFormat.Html);
       //     new InputSimulator().Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
            Thread.Sleep(200);
        }

        private void Send_Click(object sender, RoutedEventArgs e)
        {
            var sendButton = document.GetElementById("send");
            sendButton?.Focus();
            sendButton?.InvokeMember("click");
        }
    }

}
