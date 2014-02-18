using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace VkParser
{
    public partial class AuthorizationForm : Form
    {
			private string startUrl = "http://api.vkontakte.ru/oauth/authorize?client_id=4093412&scope=4&redirect_uri=http://api.vkontakte.ru/blank.html&display=popup&response_type=token&hash=0";

        public AuthorizationForm()
        {
            InitializeComponent();
        }

        private void AuthorizationForm_Load(object sender, EventArgs e)
        {
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.Navigate(startUrl);
        }

        public void SetInfo()
        {
            try
            {
                string pattern = @"(?<=access_token=)(.*)(?=&expires_in)";
                Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
                Match m = r.Match(webBrowser.Document.Url.ToString());
                UserInfo.Acces_token = m.Value;

                pattern = @"(?<=user_id=)(.*)";
                r = new Regex(pattern, RegexOptions.IgnoreCase);
                m = r.Match(webBrowser.Document.Url.ToString());
                UserInfo.User_id = m.Value;
            }
            catch
            { 
                MessageBox.Show("Ошибка подключения"); 
            }
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
					SetInfo();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }
    }
}
