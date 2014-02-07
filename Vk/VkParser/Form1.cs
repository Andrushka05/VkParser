using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml;
using System.Threading;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

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
		private Thread threadProd;
		private Label lb_state1;
		List<Comment> arrResult;
		private List<Album> albums;
		private string groupId = "";

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
					GetPhotosText(Groups[i], AlbumsId);
				}
				if (checkBox5.Checked)
				{
					ShowMessage("Группа: " + Groups[i] + " cбор данных с подписи фото");
					GetPhotosAlbum(Groups[i], ref AlbumsId);
					GetPhotosComments(Groups[i], AlbumsId);
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
					if (Comments[i].UserId != null && Comments[i].UserId.Length > 0 && Convert.ToInt64(Comments[i].UserId) < 0)
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
		private bool PhoneContains(string str)
		{
			str=str.Replace("#", "");
				string url = @"((([A-Za-z]{3,9}:(?:\/\/)?)(?:[-;:&=\+\$,\w]+@)?[A-Za-z0-9.-]+|(?:www.|[-;:&=\+\$,\w]+@)[A-Za-z0-9.-]+)((?:\/[\+~%\/.\w-_]*)?\??(?:[-\+=&;%@.\w_]*)#?(?:[\w]*))?)";
				var ms = Regex.Matches(str, url).Cast<Match>().Select(m => m.Value).ToList();
				foreach (var m in ms)
				{
					str=str.Replace(m,"");
				}
				if (str.Contains("vk.com"))
				{
					str = str.Replace("vk.com", " www.vk.com");
					ms = Regex.Matches(str, url).Cast<Match>().Select(m => m.Value).ToList();
					foreach (var m in ms)
					{
						str = str.Replace(m, "");
					}
				}
			string pattern = @"(^((8|\+7|7|\+3|3|9|0)(\s|\d|\-)*)|(?<=\W)(8|\+7|7|\+3|3|9|0)(\s|\d|\-)*)";
			string pat2 = @"(\d{1}\D{,2}\d{3}\D{,2}\d{3}\D{,2}\d{2}\D{,2}\d{2})";
			string pat3 =
					@"(?:8(?:(?:21|22|23|24|51|52|53|54|55)|(?:15\d\d))?|\+?7)?(?:(?:3[04589]|4[012789]|8[^89\D]|9\d)\d)?\d{7}";
			while (true)
			{
				Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
				Match m = r.Match(str);
				Match m2 = Regex.Match(str, pat2);
				Match m3 = Regex.Match(str, pat3);
				if (m.Value != "")
				{
					if (m.Value.Length < 9)
					{
						str = str.Replace(m.Value, "");
					}
					else
					{
						string res3 = GetDigits(m.Value);
						if (res3.Length >= 9 && res3.Length < 14 && res3[0].ToString() != "0")
						{
							return true;
						}
						else
						{
							str = str.Replace(m.Value, "");
						}
					}
				}
				else
					break;
			}
			return false;
		}
		private List<string> PhoneSearch(string str)
		{
			var res = new List<string>();
			string pattern = @"(^((8|\+7|7|\+3|3|9|0)(\s|\d|\-)*)|(?<=\W)(8|\+7|7|\+3|3|9|0)(\s|\d|\-)*)";
			string pat2 = @"((8|\+7)[\- ]?)?(\(?\d{3}\)?[\- ]?)?[\d\- ]{7,10}";
			string pat3 =
					@"(?:8(?:(?:21|22|23|24|51|52|53|54|55)|(?:15\d\d))?|\+?7)?(?:(?:3[04589]|4[012789]|8[^89\D]|9\d)\d)?\d{7}";
			while (true)
			{
				Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
				Match m = r.Match(str);
				Match m2 = Regex.Match(str, pat2);
				Match m3 = Regex.Match(str, pat3);
				if (m.Value != "")
				{
					if (m.Value.Length < 9)
					{
						str = str.Replace(m.Value, "");
					}
					else
					{
						string res3 = GetDigits(m.Value);
						if (res3.Length >= 9 && res3.Length < 14 && res3[0].ToString() != "0")
						{
							res.Add(res3);
							str = str.Replace(m.Value, "");
						}
						else
						{
							str = str.Replace(m.Value, "");
						}
					}
				}
				else
					break;
			}
			return res;
		}
		private List<string> PhoneSearch(Comment obj)
		{
			string str = obj.CommentText;
			return PhoneSearch(str);
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

		private void GetPhotosText(string group_id, List<string> AlbumsId)
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
					doc = new XmlDocument();
					doc.LoadXml(response);
					photoId = doc.GetElementsByTagName("pid");
					text = doc.GetElementsByTagName("text");
					if (photoId.Count != 0)
					{
						
						for (var i = 0; i < photoId.Count; i++)
						{
							if (!string.IsNullOrEmpty(text[i].InnerText))
							{
								string commentLink = "http://vk.com/photo-" + group_id + "_" + photoId[i].InnerText;
								Comments.Add(new Comment("", commentLink, text[i].InnerText));
							}
						}
						offset += 100;
						Thread.Sleep(350);
					}
					else
					{
						Thread.Sleep(350);
						break;
					}
				}
			}
		}

		private void GetPhotosComments(string group_id, List<string> aId)
		{

			XmlDocument doc;
			XmlNodeList fromId;
			XmlNodeList photoId;
			XmlNodeList text;
			for (var i = 0; i < aId.Count; i++)
			{
				int offset = 0;
				while (true)
				{
					string response = Request("https://api.vk.com/method/photos.getAllComments.xml?owner_id=-" + group_id + "&aid=" + aId[i] + "&offset=" +
											offset + "&count=100&access_token=" + UserInfo.Acces_token);
					//if (string.IsNullOrEmpty(response))
					//{
					//	Thread.Sleep(500);
					//	response = Request("https://api.vk.com/method/photos.getgetAllComments.xml?owner_id=-" + group_id + "&photo_id=" + photosId[i].id + "&offset=" +
					//						offset + "&count=100&access_token=" + UserInfo.Acces_token);
					//	if (string.IsNullOrEmpty(response))
					//		break;
					//}
					doc = new XmlDocument();
					doc.LoadXml(response);
					fromId = doc.GetElementsByTagName("from_id");
					photoId = doc.GetElementsByTagName("pid");
					text = doc.GetElementsByTagName("message");
					
					if (fromId.Count != 0)
					{
						for (int j = 0; j < fromId.Count; j++)
						{
							string commentLink = "http://vk.com/photo-" + group_id + "_" + photoId[i].InnerText;
							if (!string.IsNullOrEmpty(text[j].InnerText.Trim()) && PhoneContains(text[j].InnerText))
								Comments.Add(new Comment(fromId[j].InnerText, commentLink, text[j].InnerText));
						}
						
						offset += 100;
						Thread.Sleep(150);
					}
					else
					{
						Thread.Sleep(350);
						break;
					}
				}
			}
		}

		private void GetTopicsId(string group_id, ref List<string> TopicsId)
		{
			int offset = 0;
			while (true)
			{
				string response =
						Request("https://api.vk.com/method/board.getTopics.xml?group_id=" + group_id +
										"&offset=" + offset + "&count=100&access_token=" + UserInfo.Acces_token);
				XmlDocument doc = new XmlDocument();
				doc.LoadXml(response);
				XmlNodeList id = doc.GetElementsByTagName("tid");
				if (id.Count != 0)
				{
					foreach (XmlNode n in id)
					{
						TopicsId.Add(n.InnerText);
					}
					offset += 100;
					Thread.Sleep(350);
				}
				else
					break;
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

		private void btnAlbum_Click(object sender, EventArgs e)
		{
			Control.CheckForIllegalCrossThreadCalls = false;
			treeView1.Nodes.Clear();
			btnAlbum.Enabled = false;
			btnParser.Visible = false;
			this.Update();
			threadProd = new Thread(GetNameAlbum);
			threadProd.Start();
		}

		private void GetNameAlbum()
		{
			try
			{
				string pattern = @"(?<=com\/).*";
				albums = new List<Album>();
				Thread.Sleep(1000);
				Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
				Match m = r.Match(tbLink.Text);
				if (m.Value != "")
				{
					XmlDocument doc = new XmlDocument();
					string response =
							Request("https://api.vk.com/method/groups.getById.xml?group_id=" + m.Value + "&access_token=" +
											UserInfo.Acces_token);
					doc.LoadXml(response);
					XmlNodeList n = doc.GetElementsByTagName("gid");
					groupId = n[0].InnerText;
					int offset = 0;
					while (true)
					{
						string response2 =
								Request("https://api.vk.com/method/photos.getAlbums.xml?owner_id=-" + groupId + "&offset=" +
												offset +
												"&count=100&filter=all&access_token=" + UserInfo.Acces_token);
						Thread.Sleep(350);
						XmlDocument doc2 = new XmlDocument();
						doc2.LoadXml(response2);
						XmlNodeList ids = doc2.GetElementsByTagName("aid");
						XmlNodeList names = doc2.GetElementsByTagName("title");
						if (ids.Count != 0)
						{
							for (var i = 0; i < ids.Count; i++)
							{
								albums.Add(new Album() { Id = ids[i].InnerText, Name = names[i].InnerText });
							}
							offset += 100;
						}
						else
							break;
					}
					if (albums.Any())
					{
						var nodes = new TreeNode();
						nodes.Name = "Альбомы";
						foreach (var album in albums.OrderBy(x => x.Name))
						{
							nodes.Nodes.Add(album.Name);
						}
						this.Invoke(new Action(() =>
						{
							treeView1.Nodes.Add(nodes);

						}));
					}
				}
			}
			catch (Exception e)
			{
				MessageBox.Show("Ошибка: " + e);
			}
			this.Invoke(new Action(() => { btnAlbum.Enabled = true; }));
			this.Invoke(new Action(() => { btnParser.Visible = true; }));
		}

		private void btnParser_Click(object sender, EventArgs e)
		{
			Control.CheckForIllegalCrossThreadCalls = false;
			btnAlbum.Enabled = false;
			btnParser.Enabled = false;
			this.Update();
			threadProd = new Thread(GetItemsOfAlbum);
			threadProd.Start();
		}

		private void GetItemsOfAlbum()
		{
			var parsAlbum = new List<Album>();

			foreach (TreeNode aNode in treeView1.Nodes)
			{
				foreach (TreeNode n in aNode.Nodes)
					if (n.Checked)
					{
						parsAlbum.Add(albums.FirstOrDefault(x => x.Name == n.Text));
					}
			}
			var ex2003 = rb2003.Checked;
			foreach (var album in parsAlbum)
			{
				var products = new List<Product>();
				int offset = 0;
				while (true)
				{
					string response =
							Request("https://api.vk.com/method/photos.get.xml?owner_id=-" + groupId + "&album_id=" +
											album.Id + "&offset=" + offset + "&count=100&access_token=" + UserInfo.Acces_token);
					Thread.Sleep(350);
					var doc = new XmlDocument();
					doc.LoadXml(response);
					var photoId = doc.GetElementsByTagName("pid");
					var photos = doc.GetElementsByTagName("src_xxbig");
					if (photos.Count != photoId.Count)
					{
						photos = doc.GetElementsByTagName("src_xbig");
						if (photos.Count != photoId.Count)
							photos = doc.GetElementsByTagName("src_big");
					}
					var text = doc.GetElementsByTagName("text");
					if (photoId.Count != 0)
					{
						for (var i = 0; i < photoId.Count; i++)
						{
							products.Add(new Product() { Id = photoId[i].InnerText, Text = text[i].InnerText.Replace("<br>", "\r\n"), Src = photos[i].InnerText });
						}
						offset += 100;
						Thread.Sleep(350);
					}
					else
						break;
				}
				//save to file
				var path = album.Name.Replace(" ", "_").Replace("*", "x").Replace("\"", "_").Replace("/", "_").Replace("<", "_").Replace("?", "_").Replace(">", "_").Replace("\\", "_").Replace(":", "_").Replace("|", "_");
				if (!ex2003)
					SaveExcel2007(products, Environment.CurrentDirectory + @"\" + path + "_" + DateTime.Now.Hour + "" + DateTime.Now.Minute + "" + DateTime.Now.Second + ".xlsx");
				else
					SaveExcel2003(products, Environment.CurrentDirectory + @"\" + path + "_" + DateTime.Now.Hour + "" + DateTime.Now.Minute + "" + DateTime.Now.Second + ".xls");
			}


			this.Invoke(new Action(() => { btnAlbum.Enabled = true; }));
			this.Invoke(new Action(() => { btnParser.Enabled = true; }));
		}

		public static void SaveExcel2007(IEnumerable<Product> list, string path)
		{

			using (var p = new ExcelPackage(File.Exists(path) ? new FileInfo(path) : null))
			{
				p.Workbook.Properties.Title = "VK";

				//Create a sheet
				//p.Workbook.Worksheets.Add("Sample WorkSheet");
				ExcelWorksheet ws = null;
				int colIndex = 1;
				int rowIndex = 1;
				if (p.Workbook.Worksheets.Count == 0)
				{
					ws = p.Workbook.Worksheets.Add("Sample WorkSheet");
					ws.Name = "Vk"; //Setting Sheet's name
					ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
					ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

					colIndex = 1;

					var cell = ws.Cells[rowIndex, colIndex];
					var cel2 = ws.Cells[rowIndex + 1, colIndex];
					var cel3 = ws.Cells[rowIndex, colIndex + 1];
					var cel4 = ws.Cells[rowIndex + 1, colIndex + 1];
					var cel5 = ws.Cells[rowIndex, colIndex + 2];
					var cel6 = ws.Cells[rowIndex + 1, colIndex + 2];
					//Setting the background color of header cells to Gray
					var fill = cel2.Style.Fill;
					//fill.PatternType = ExcelFillStyle.Solid;
					//fill.BackgroundColor.SetColor(Color.Gray);
					cel2.Style.Font.Bold = true;
					cel4.Style.Font.Bold = true;
					cel6.Style.Font.Bold = true;

					//Setting Top/left,right/bottom borders.
					var border = cel2.Style.Border = cel4.Style.Border = cel6.Style.Border;
					border.Bottom.Style =
							border.Top.Style =
							border.Left.Style =
							border.Right.Style = ExcelBorderStyle.Thin;
					cell.Value = "id";
					cel2.Value = "id";
					cel3.Value = "text";
					cel4.Value = "Описание";
					cel5.Value = "src";
					cel6.Value = "Фото";

				}
				else
					ws = p.Workbook.Worksheets.First();

				rowIndex = ws.Dimension.End.Row;
				foreach (var l in list) // Adding Data into rows
				{
					rowIndex++;
					var cell = ws.Cells[rowIndex, colIndex];
					var cel2 = ws.Cells[rowIndex, colIndex + 1];
					var cel3 = ws.Cells[rowIndex, colIndex + 2];

					var border = cell.Style.Border = cel2.Style.Border = cel3.Style.Border;
					border.Left.Style =
							border.Right.Style = ExcelBorderStyle.Thin;
					cell.Value = l.Id;
					cel2.Value = l.Text;
					cel3.Value = l.Src;
				}

				Byte[] bin = p.GetAsByteArray();
				File.WriteAllBytes(path, bin);
			}
		}

		public static void SaveExcel2003(IEnumerable<Product> list, string path)
		{
			IWorkbook wb = new HSSFWorkbook();


			Dictionary<String, ICellStyle> styles = createStyles(wb);

			ISheet sheet = wb.CreateSheet("Item");

			//turn off gridlines
			sheet.DisplayGridlines = (false);
			sheet.IsPrintGridlines = (false);
			sheet.FitToPage = (true);
			sheet.HorizontallyCenter = (true);
			IPrintSetup printSetup = sheet.PrintSetup;
			printSetup.Landscape = (true);
			sheet.Autobreaks = (true);
			printSetup.FitHeight = ((short)1);
			printSetup.FitWidth = ((short)1);

			//the header row: centered text in 48pt font
			IRow headerRow = sheet.CreateRow(1);
			//headerRow.HeightInPoints = (12.75f);
			ICell cell = headerRow.CreateCell(0);
			cell.SetCellValue("id");
			ICell cel2 = headerRow.CreateCell(1);
			cel2.SetCellValue("text");
			ICell cel3 = headerRow.CreateCell(2);
			cel3.SetCellValue("src");

			IRow headerRow1 = sheet.CreateRow(2);
			//headerRow.HeightInPoints = (12.75f);
			ICell cel4 = headerRow1.CreateCell(0);
			cell.SetCellValue("id");
			ICell cel5 = headerRow1.CreateCell(1);
			cel2.SetCellValue("Описание");
			ICell cel6 = headerRow1.CreateCell(2);
			cel3.SetCellValue("Фото");

			//freeze the first row
			//sheet.CreateFreezePane(0, 1);

			IRow row;
			//ICell cell;
			int rownum = 1;
			foreach (var l in list) // Adding Data into rows
			{
				rownum++;
				row = sheet.CreateRow(rownum);
				var c1 = row.CreateCell(0);
				var c2 = row.CreateCell(1);
				var c3 = row.CreateCell(2);
				c1.CellStyle = styles["cell_normal"];
				c2.CellStyle = styles["cell_normal"];
				c3.CellStyle = styles["cell_normal"];
				c1.SetCellValue(l.Id);
				c2.SetCellValue(l.Text);
				c3.SetCellValue(l.Src);
			}

			// Write the output to a file
			using (FileStream out1 = new FileStream(path, FileMode.Create))
			{
				wb.Write(out1);
				out1.Close();
			}
		}
		private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
		{
			// The code only executes if the user caused the checked state to change.
			if (e.Action != TreeViewAction.Unknown)
			{
				if (e.Node.Nodes.Count > 0)
				{
					/* Calls the CheckAllChildNodes method, passing in the current 
					Checked value of the TreeNode whose checked state changed. */
					this.CheckAllChildNodes(e.Node, e.Node.Checked);
				}
			}
		}
		private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
		{
			foreach (TreeNode node in treeNode.Nodes)
			{
				node.Checked = nodeChecked;
				if (node.Nodes.Count > 0)
				{
					// If the current node has child nodes, call the CheckAllChildsNodes method recursively.
					this.CheckAllChildNodes(node, nodeChecked);
				}
			}
		}
		private static Dictionary<String, ICellStyle> createStyles(IWorkbook wb)
		{
			Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();
			IDataFormat df = wb.CreateDataFormat();

			ICellStyle style;
			IFont headerFont = wb.CreateFont();
			headerFont.Boldweight = (short)(FontBoldWeight.Bold);
			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
			style.FillPattern = FillPattern.SolidForeground;
			style.SetFont(headerFont);
			styles.Add("header", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
			style.FillPattern = FillPattern.SolidForeground;
			style.SetFont(headerFont);
			style.DataFormat = (df.GetFormat("d-mmm"));
			styles.Add("header_date", style);

			IFont font1 = wb.CreateFont();
			font1.Boldweight = (short)(FontBoldWeight.Bold);
			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.SetFont(font1);
			styles.Add("cell_b", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.SetFont(font1);
			styles.Add("cell_b_centered", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.SetFont(font1);
			style.DataFormat = (df.GetFormat("d-mmm"));
			styles.Add("cell_b_date", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.SetFont(font1);
			style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
			style.FillPattern = FillPattern.SolidForeground;
			style.DataFormat = (df.GetFormat("d-mmm"));
			styles.Add("cell_g", style);

			IFont font2 = wb.CreateFont();
			font2.Color = (IndexedColors.Blue.Index);
			font2.Boldweight = (short)(FontBoldWeight.Bold);
			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.SetFont(font2);
			styles.Add("cell_bb", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.SetFont(font1);
			style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
			style.FillPattern = FillPattern.SolidForeground;
			style.DataFormat = (df.GetFormat("d-mmm"));
			styles.Add("cell_bg", style);

			IFont font3 = wb.CreateFont();
			font3.FontHeightInPoints = ((short)14);
			font3.Color = (IndexedColors.DarkBlue.Index);
			font3.Boldweight = (short)(FontBoldWeight.Bold);
			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.SetFont(font3);
			style.WrapText = (true);
			styles.Add("cell_h", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.WrapText = (true);
			styles.Add("cell_normal", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.WrapText = (true);
			styles.Add("cell_normal_centered", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.WrapText = (true);
			style.DataFormat = (df.GetFormat("d-mmm"));
			styles.Add("cell_normal_date", style);

			style = CreateBorderedStyle(wb);
			style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			style.Indention = ((short)1);
			style.WrapText = (true);
			styles.Add("cell_indented", style);

			style = CreateBorderedStyle(wb);
			style.FillForegroundColor = (IndexedColors.Blue.Index);
			style.FillPattern = FillPattern.SolidForeground;
			styles.Add("cell_blue", style);

			return styles;
		}
		private static ICellStyle CreateBorderedStyle(IWorkbook wb)
		{
			ICellStyle style = wb.CreateCellStyle();
			style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
			style.RightBorderColor = (IndexedColors.Black.Index);
			style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
			style.BottomBorderColor = (IndexedColors.Black.Index);
			style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
			style.LeftBorderColor = (IndexedColors.Black.Index);
			style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
			style.TopBorderColor = (IndexedColors.Black.Index);
			return style;
		}
	}
}
