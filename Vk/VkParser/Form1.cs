using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Threading;
using System.Text.RegularExpressions;

namespace VkParser
{
    public partial class Form1 : Form
    {
        private bool Loaded;
        private string startUrl = "http://api.vkontakte.ru/oauth/authorize?client_id=4093412&scope=8196&redirect_uri=http://api.vkontakte.ru/blank.html&display=popup&response_type=token&hash=0";
        private List<string> Groups;
        private List<Comment> Comments;
        private WebBrowser webBrowser;
        private Thread tread;
        private Label lb_state1;
        List<Comment> arrResult;

        public Form1()
        {
            InitializeComponent();

            //Comment cm = new Comment("asd", "asd", "asdas das asd sdfa 140<a>9373235707");//9373235707
            //string df = PhoneSearch(cm);
            webBrowser = new WebBrowser();
            webBrowser.ScriptErrorsSuppressed = true;
            Loaded = false;
            webBrowser.Navigate(startUrl);
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);
            while (!Loaded)
            {
                Application.DoEvents();
            }
            SetInfo();
            if (UserInfo.Acces_token == "")
            {
                AuthorizationForm auth = new AuthorizationForm();
                auth.ShowDialog();
            }       
         }

        private void bt_start_Click(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            tread = new Thread(Main);
            tread.Start();
        }

        private void Main()
        {
            //DrawLb();
            Groups = new List<string>();
            Comments = new List<Comment>();
            ShowMessage("Чтение из файла");
            if (tb_path.Text != "")
            {
                Groups = GetGroup();
                if (Groups == null)
                {
                    ShowMessage("Ошибка чтения из файла");
                    return;
                }
            }
            else
            {
                ShowMessage("Файл не выбран");
                return;
            }
            for (int i = 0; i < Groups.Count; i++)
            {
                List<string> PostsId = new List<string>();
                List<string> TopicsId = new List<string>();
                List<string> AlbumsId = new List<string>();
                if (checkBox4.Checked)
                {
                    ShowMessage("Группа: " + Groups[i] + " cбор данных со стены");
                    GetWall(Groups[i], ref PostsId);
                }
                if (checkBox1.Checked)
                {
                    ShowMessage("Группа: " + Groups[i] + " cбор данных из комментариев");
                    GetWallComments(Groups[i], PostsId);
                }
                if (checkBox2.Checked)
                {
                    ShowMessage("Группа: " + Groups[i] + " cбор данных из обсуждений");
                    GetTopicsId(Groups[i], ref TopicsId);
                    GetTopicComments(Groups[i], TopicsId);
                }
                if (checkBox3.Checked)
                {
                    ShowMessage("Группа: " + Groups[i] + " cбор данных с изображений");
                    GetPhotosAlbum(Groups[i], ref AlbumsId);
                    var photoIds = GetPhotosId(Groups[i], AlbumsId);
                    GetPhotosComments(Groups[i],photoIds);
                }
            }
            ShowMessage("Обработка данных");
            for (int i = 0; i < Comments.Count; i++)
            {
                var res = PhoneSearch(Comments[i]);
                if (res.Any())
                {
                    Comments[i].UserPhone = res;
                }
            }
            SaveCsv();
            ShowMessage("Готово");
        }

        private string GetDigits(string s)
        {
            try
            {
                string result = "";
                for (int i = 0; i < s.Length; i++)
                {
                    if (char.IsDigit(s[i]))
                    {
                        result += s[i];
                    }
                }
                return result;
            }
            catch
            { 
                return null; 
            }
        }

        private void SetInfo()
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

        private List<string> GetGroup()
        {
            try
            {
                List<string> groups = new List<string>();
                StreamReader stream = new StreamReader(tb_path.Text);
                string t = "";
                while (t != null)
                {
                    t = stream.ReadLine();
                    if (t != null)
                    {
                        groups.Add(t);
                    }
                }
                string pattern = @"(?<=com\/).*";
                for (int i = 0; i < groups.Count; i++)
                {
                    Thread.Sleep(1000);
                    Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
                    Match m = r.Match(groups[i]);
                    if (m.Value != "")
                    {
                        XmlDocument doc = new XmlDocument();
                        string response = Request("https://api.vk.com/method/groups.getById.xml?group_id=" + m.Value + "&access_token=" + UserInfo.Acces_token);
                        doc.LoadXml(response);
                        XmlNodeList n = doc.GetElementsByTagName("gid");
                        groups[i] = n[0].InnerText;
                    }
                }
                return groups;
            }
            catch
            {
                return null;
            }
        }
        /// <summary>.
        /// full
        /// </summary>
        //private void SaveCsv()
        //{
        //    if (Comments.Count != 0)
        //    {
        //        ShowMessage("Получение данных о пользователе");
        //        List<Comment> UsersWithPhone = new List<Comment>();
        //        for (int i = 0; i < Comments.Count; i++)
        //        {
        //            if (Comments[i].UserPhone != null)
        //            {
        //                UsersWithPhone.Add(Comments[i]);
        //            }
        //        }
        //        Comments = UsersWithPhone;
        //        for (int i = 0; i < Comments.Count; i++)
        //        {
        //            string p = Comments[i].UserId;
        //            bool f = false;
        //            for (int j = 0; j < Comments.Count; j++)
        //            {
        //                if (!f)
        //                {
        //                    if (p == Comments[j].UserId)
        //                    {
        //                        f = true;
        //                    }
        //                }
        //                else
        //                {
        //                    if (p == Comments[j].UserId)
        //                    {
        //                        Comments[j].UserPhone = null;
        //                    }
        //                }
        //            }
        //        }
        //        UsersWithPhone = new List<Comment>();
        //        for (int i = 0; i < Comments.Count; i++)
        //        {
        //            if (Comments[i].UserPhone != null)
        //            {
        //                UsersWithPhone.Add(Comments[i]);
        //            }
        //        }
        //        Comments = UsersWithPhone;
        //        List<Comment> Admin = new List<Comment>();
        //        for (int i = 0; i < Comments.Count; i++)
        //        {
        //            if (Convert.ToInt64(Comments[i].UserId) < 0)
        //            {
        //                Comments[i].UserName = "admin";
        //                Admin.Add(Comments[i]);
        //                Comments.RemoveAt(i);
        //            }
        //        }
        //        if (Comments.Count > 250)
        //        {
        //            int arrCount = Comments.Count;
        //            int r = arrCount / 250;
        //            int sum = 0;
        //            arrResult = new List<Comment>();
        //            List<int> arrVal = new List<int>();
        //            for (int j = 0; j < r; j++)
        //            {
        //                arrVal.Add(250);
        //                sum += 250;
        //            }
        //            arrCount -= sum;
        //            arrVal.Add(arrCount);

        //            int count = 0;
        //            int q = 0;
        //            int offset = 0;
        //            List<Comment> temp = new List<Comment>();
        //            while (true)
        //            {
        //                if (offset == arrVal[q])
        //                {
        //                    GetUser(ref temp);
        //                    for (int y = 0; y < temp.Count; y++)
        //                    {
        //                        arrResult.Add(temp[y]);
        //                    }
        //                    temp = new List<Comment>();
        //                    q++;
        //                    offset = 0;
        //                    if (q == arrVal.Count)
        //                    {
        //                        break;
        //                    }
        //                    temp.Add(Comments[count]);
        //                }
        //                else
        //                {
        //                    if (count == Comments.Count)
        //                        break;
        //                    temp.Add(Comments[count]);
        //                }
        //                count++;
        //                offset++;
        //            }
        //        }
        //        else
        //        {
        //            arrResult = new List<Comment>();
        //            arrResult = Comments;
        //            GetUser(ref arrResult);
        //        }
        //        try
        //        {
        //            string date = DateTime.Now.ToString("HH:mm:ss").Replace(":", ".");
        //            StreamWriter csv = new StreamWriter(date + ".csv", true, Encoding.Unicode);
        //            for (int i = 0; i < arrResult.Count; i++)
        //            {
        //                if (arrResult[i].UserPhone != null)
        //                {
        //                    csv.WriteLine(arrResult[i].UserName + ";" + arrResult[i].UserPhone + ";http://vk.com/id" + arrResult[i].UserId + ";" + arrResult[i].CommentLink);
        //                }
        //            }
        //            csv.Close();
        //        }
        //        catch
        //        {
        //            ShowMessage("Ошибка сохранения");
        //        }
        //    }
        //}

        /// <summary>
        /// demo
        /// </summary>
        private void SaveCsv()
        {
            if (Comments.Count != 0)
            {
                ShowMessage("Получение данных о пользователе");
                List<Comment> UsersWithPhone = new List<Comment>();
                for (int i = 0; i < Comments.Count; i++)
                {
                    if (Comments[i].UserPhone != null)
                    {
                        UsersWithPhone.Add(Comments[i]);
                    }
                }
                Comments = UsersWithPhone;
                for (int i = 0; i < Comments.Count; i++)
                {
                    string p = Comments[i].UserId;
                    bool f = false;
                    for (int j = 0; j < Comments.Count; j++)
                    {
                        if (!f)
                        {
                            if (p == Comments[j].UserId)
                            {
                                f = true;
                            }
                        }
                        else
                        {
                            if (p == Comments[j].UserId)
                            {
                                Comments[j].UserPhone = null;
                            }
                        }
                    }
                }
                UsersWithPhone = new List<Comment>();
                for (int i = 0; i < Comments.Count; i++)
                {
                    if (Comments[i].UserPhone != null)
                    {
                        UsersWithPhone.Add(Comments[i]);
                    }
                }
                Comments = UsersWithPhone;
                List<Comment> Admin = new List<Comment>();
                for (int i = 0; i < Comments.Count; i++)
                {
									if (Comments[i].UserId!=null&&Convert.ToInt64(Comments[i].UserId) < 0)
                    {
                        Comments[i].UserName = "admin";
                        Admin.Add(Comments[i]);
                        Comments.RemoveAt(i);
                    }
                }

                if (Comments.Count > 50)
                {
                    arrResult = new List<Comment>();
                    for (int q = 0; q < 50; q++)
                    {
                        arrResult.Add(Comments[q]);
                    }
                    GetUser(ref arrResult);
                }
                else
                {
                    arrResult = new List<Comment>();
                    arrResult = Comments;
                    GetUser(ref arrResult);
                }

                try
                {
                    string date = DateTime.Now.ToString("HH:mm:ss").Replace(":", ".");
                    StreamWriter csv = new StreamWriter(date + ".csv", true, Encoding.Unicode);
                    for (int i = 0; i < arrResult.Count; i++)
                    {
                        if (arrResult[i].UserPhone != null)
                        {
                            var phone = string.Join("; ", arrResult[i].UserPhone);
                            csv.WriteLine(arrResult[i].UserName + ";" + phone + ";http://vk.com/id" + arrResult[i].UserId + ";" + arrResult[i].CommentLink);
                        }
                    }
                    csv.Close();
                }
                catch
                {
                    ShowMessage("Ошибка сохранения");
                }
            }
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            Loaded = true;
        }

        private List<string> PhoneSearch(Comment obj)
        {
            string pattern = @"(^((8|\+7|7|\+3|3|9|0)(\s|\d|\-)*)|(?<=\W)(8|\+7|7|\+3|3|9|0)(\s|\d|\-)*)";
					var pattern2=@"((8|\+7)[\- ]?)?(\(?\d{3}\)?[\- ]?)?[\d\- ]{7,10}";
            string str = obj.CommentText;
						if (str.Contains("http"))
							str = str.Remove(str.IndexOf("http"), str.IndexOf(" ", str.IndexOf("http")) - str.IndexOf("http"));
					if(str.Contains("id"))
						str = str.Remove(str.IndexOf("id"), str.IndexOf("]", str.IndexOf("id")) - str.IndexOf("id"));
            while (true)
            {
                Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
								Regex r2 = new Regex(pattern2, RegexOptions.IgnoreCase);
                var ms = r.Matches(str).Cast<Match>().Select(m => m.Value).ToList();
								var ms2 = r2.Matches(str).Cast<Match>().Select(m => m.Value).ToList();
                var res = (from m in ms where m != "" where m.Length >= 9 select GetDigits(m) into num where num.Length >= 6&&num.Length<11 select num).ToList();
								var res2 = (from m in ms2 where m != "" where m.Length >= 7 select GetDigits(m) into num where num.Length >= 6 && num.Length < 11 select num).ToList();
								if (ms2.Count > 0)
								{
									var ter = res2;
								}
                return res;
            }
        }

        private void GetUser(ref List<Comment> arr)
        {
            try
            {
                string user_ids = "";
                for (int i = 0; i < arr.Count; i++)
                {
                    user_ids += arr[i].UserId + ",";
                }
                user_ids = user_ids.Remove(user_ids.Length - 1);
                string response = Request("https://api.vk.com/method/users.get.xml?user_ids=" + user_ids + "&access_token=" + UserInfo.Acces_token);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(response);
                XmlNodeList fName = doc.GetElementsByTagName("first_name");
                XmlNodeList sName = doc.GetElementsByTagName("last_name");
                for (int j = 0; j < fName.Count; j++)
                {
                    arr[j].UserName = fName[j].InnerText + " " + sName[j].InnerText;
                }
            }
            catch { }
        }

        private void GetPhotosAlbum(string group_id, ref List<string> AlbumsId)
        {
            List<string> FromId = new List<string>();
            List<string> Text = new List<string>();
            int offset = 0;
            while (true)
            {
                string response =
                    Request("https://api.vk.com/method/photos.getAlbums.xml?owner_id=-" + group_id + "&offset=" + offset +
                            "&count=100&filter=all&access_token=" + UserInfo.Acces_token);
                Thread.Sleep(350);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(response);
                XmlNodeList id = doc.GetElementsByTagName("aid");
                if (id.Count != 0)
                {
                    foreach (XmlNode n in id)
                    {
                        AlbumsId.Add(n.InnerText);
                    }
                    offset += 100;
                }
                else
                    break;
            }
        }

        private List<Photo> GetPhotosId(string group_id, List<string> AlbumsId)
        {
            var res = new List<Photo>();
            XmlDocument doc;
            XmlNodeList photoId;
					XmlNodeList text;
            foreach (var alb in AlbumsId)
            {
                int offset = 0;
                while (true)
                {
                    string response =
                        Request("https://api.vk.com/method/photos.get.xml?owner_id=-" + group_id + "&album_id=" +
                                alb + "&offset=" + offset + "&count=100&access_token=" + UserInfo.Acces_token);
                    Thread.Sleep(350);
                    doc = new XmlDocument();
                    doc.LoadXml(response);
                    photoId = doc.GetElementsByTagName("pid");
									text = doc.GetElementsByTagName("text");
                    if (photoId.Count != 0)
                    {
											for (var i = 0; i < photoId.Count; i++)
											{
												res.Add(new Photo() { id = photoId[i].InnerText, text = text[i].InnerText});
											}
												offset += 100;
                        Thread.Sleep(350);
                    }
                    else
                        break;
                }
            }
            return res;
        }

				private void GetPhotosComments(string group_id, List<Photo> photosId)
        {
            
            XmlDocument doc;
            XmlNodeList fromId;
            XmlNodeList photoId;
            XmlNodeList text;
						for (var i = 0; i < photosId.Count; i++)
            {
                int offset = 0;
                while (true)
                {
                    string response =
												Request("https://api.vk.com/method/photos.getComments.xml?owner_id=-" + group_id + "&photo_id=" + photosId[i].id + "&offset=" +
                                offset + "&count=100&access_token=" + UserInfo.Acces_token);
                    Thread.Sleep(350);
                    doc = new XmlDocument();
                    doc.LoadXml(response);
                    fromId = doc.GetElementsByTagName("from_id");
                    photoId = doc.GetElementsByTagName("pid");
                    text = doc.GetElementsByTagName("message");
										string commentLink = "http://vk.com/photo-" + group_id + "_" + photosId[i].id;
										if (fromId.Count != 0)
										{
											for (int j = 0; j < fromId.Count; j++)
											{
												if(j==0)
													Comments.Add(new Comment(fromId[j].InnerText, commentLink, text[j].InnerText + "\r\n" + photosId[i].text));
												else
													Comments.Add(new Comment(fromId[j].InnerText, commentLink, text[j].InnerText));
											}
											offset += 100;
											Thread.Sleep(350);
										}
										else
										{
											Comments.Add(new Comment("", commentLink, photosId[i].text));
											break;
										}
                }
            }
        }

        private void GetTopicsId(string group_id, ref List<string> TopicsId)
        {
            string response = Request("https://api.vk.com/method/board.getTopics.xml?group_id=" + group_id + "&count=100&access_token=" + UserInfo.Acces_token);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response);
            XmlNodeList id = doc.GetElementsByTagName("tid");
            foreach (XmlNode n in id)
            {
                TopicsId.Add(n.InnerText);
            }
        }

        private void GetTopicComments(string group_id, List<string> TopicsId)
        {
            XmlDocument doc;
            XmlNodeList fromId;
            XmlNodeList commentId;
            XmlNodeList text;
            for (int i = 0; i < TopicsId.Count; i++)
            {
                int offset = 0;
                while (true)
                {
                    string response = Request("https://api.vk.com/method/board.getComments.xml?group_id=" + group_id + "&topic_id=" + TopicsId[i] + "&offset=" + offset + "&count=100&access_token=" + UserInfo.Acces_token);
                    doc = new XmlDocument();
                    doc.LoadXml(response);
                    fromId = doc.GetElementsByTagName("from_id");
                    commentId = doc.GetElementsByTagName("id");
                    text = doc.GetElementsByTagName("text");
                    if (fromId.Count != 0)
                    {
                        for (int j = 0; j < fromId.Count; j++)
                        {
                            string commentLink = "http://vk.com/topic-" + group_id + "_" + TopicsId[i] + "?post=" + commentId[j].InnerText;
                            Comments.Add(new Comment(fromId[j].InnerText, commentLink, text[j].InnerText));
                        }
                        offset += 100;
                        Thread.Sleep(350);
                    }
                    else
                        break;
                }
            }
        }

        private void GetWall(string group_id, ref List<string> PostsId)
        {
            List<string> FromId = new List<string>();
            List<string> Text = new List<string>();
            int offset = 0;
            while (true)
            {
                string response = Request("https://api.vk.com/method/wall.get.xml?owner_id=-" + group_id + "&offset=" + offset + "&count=100&filter=all&access_token=" + UserInfo.Acces_token);
                Thread.Sleep(350);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(response);
                XmlNodeList id = doc.GetElementsByTagName("id");
                XmlNodeList text = doc.GetElementsByTagName("text");
                XmlNodeList fromId = doc.GetElementsByTagName("from_id");
                XmlNode pattern = text[0];
                if (id.Count != 0)
                {
                    foreach (XmlNode n in text)
                    {
                        string asd = n.ParentNode.Name;
                        if (n.ParentNode.Name == "post")
                        {
                            Text.Add(n.InnerText);
                        }
                    }
                    for (int i = 0; i < id.Count; i++)
                    {
                        PostsId.Add(id[i].InnerText);
                        FromId.Add(fromId[i].InnerText);
                    }
                    offset += 100;
                }
                else
                    break;
            }
            if (FromId.Count != 0)
            {
                for (int i = 0; i < FromId.Count; i++)
                {
                    string commentLink = "https://vk.com/club" + group_id + "?w=wall-" + group_id + "_" + PostsId[i];
                    Comments.Add(new Comment(FromId[i], commentLink, Text[i]));
                }
            }
        }

        private void GetWallComments(string group_id, List<string> PostsId)
        {
            XmlDocument doc;
            XmlNodeList fromId;
            XmlNodeList text;
            XmlNodeList err_code;
            for (int i = 0; i < PostsId.Count; i++)
            {
                int offset = 0;
                while (true)
                {
                    Thread.Sleep(500);
                    string response = Request("https://api.vk.com/method/wall.getComments.xml?owner_id=-" + group_id + "&post_id=" + PostsId[i] + "&offset=" + offset + "&count=100&access_token=" + UserInfo.Acces_token);
                    doc = new XmlDocument();
                    doc.LoadXml(response);
                    string commentLink = "http://vk.com/club" + group_id + "?w=wall-" + group_id + "_" + PostsId[i];
                    fromId = doc.GetElementsByTagName("from_id");
                    text = doc.GetElementsByTagName("text");
                    err_code = doc.GetElementsByTagName("error_code");
                    if (err_code.Count != 0)
                    {
                        i--;
                    }
                    else
                    {
                        if (fromId.Count != 0)
                        {
                            for (int j = 0; j < fromId.Count; j++)
                            {
                                Comments.Add(new Comment(fromId[j].InnerText, commentLink, text[j].InnerText));
                            }
                            offset += 100;
                        }
                        else
                            break;
                    }
                }
            }
        }

        private string Request(string url)
        {
            try
            {
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.AllowAutoRedirect = false;
                httpWebRequest.Method = "GET";
                httpWebRequest.Referer = "http://google.com";
                using (var httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse())
                {
                    using (var stream = httpWebResponse.GetResponseStream())
                    {
                        using (var reader = new StreamReader(stream, Encoding.GetEncoding(httpWebResponse.CharacterSet)))
                        {
                            return reader.ReadToEnd();
                        }
                    }
                }
            }
            catch
            {
                return String.Empty;
            }
        }

        private void ShowMessage(string message)
        {
            lb_state.Text = message;
            lb_state.Update();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                tb_path.Text = ofd.FileName;
            }
        }

        private void bt_exit_Click(object sender, EventArgs e)
        {
            while (true)
            {
                Loaded = false;
                webBrowser.Navigate("http://vk.com/support?act=new");
                while (!Loaded)
                {
                    Application.DoEvents();
                }
                Loaded = false;
                HtmlElement el = webBrowser.Document.GetElementById("logout_link");
                if (el != null)
                    el.InvokeMember("click");
                while (!Loaded)
                {
                    Application.DoEvents();
                }

                Loaded = false;
                webBrowser.Navigate(startUrl);
                while (!Loaded)
                {
                    Application.DoEvents();
                }
                if (webBrowser.Document.GetElementById("install_allow") != null)
                {
                    UserInfo.Acces_token = "";
                    UserInfo.User_id = "";
                    ShowMessage("Выход выполнен");
                    break;
                }
            }
        }

        private void bt_stop_Click(object sender, EventArgs e)
        {
            tread.Abort();
            ShowMessage("Сохранение");
            SaveCsv();
            ShowMessage("Готово");
        }

        private void DrawLb()
        {
            lb_state1 = new Label();
            lb_state1.AutoSize = true;
            lb_state1.Location = new System.Drawing.Point(12, 303);
            lb_state1.Name = "lb_state1";
            lb_state1.Size = new System.Drawing.Size(42, 13);
            lb_state1.TabIndex = 1;
            lb_state1.Text = "Готово";
            Controls.Add(lb_state1);
        }
    }
}
