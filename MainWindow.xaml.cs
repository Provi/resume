using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Net;
using System.IO;
using HtmlAgilityPack;
using System.Windows.Threading;
using System.Diagnostics;
using System.Windows.Media;

namespace Statistika
{
	public partial class MainWindow : Window
	{
		private CookieCollection Cooks = new CookieCollection();
		string date1 = DateTime.Now.Date.ToShortDateString();
		string date2 = DateTime.Now.Date.ToShortDateString();

		string numCalls_name;
		string numCalls;
		string sumCost_name;
		string sumCost;
		string sumDuration_name;
		string sumDuration;
		string trafIn_name;
		string trafIn;
		string trafOut_name;
		string trafOut;
		string sumTraf_name;
		string sumTraf;

		private System.Windows.Forms.NotifyIcon m_notifyIcon;

		public MainWindow()
		{
			try
			{
				InitializeComponent();
				ComboBox();
				ComboBox_Timer();
				dispatcherTimer();
				btn_refresh_Click(null, null);
				Scheduler();
			}
			catch
			{
				MessageBox.Show("Нет подключения к Internet");
				btn_refresh.IsEnabled = false;
			}
		}

		#region Нажатие кнопки "Обновить"
		private void btn_refresh_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				try { RequestTelefonia_Post(); }
				catch { MessageBox.Show("Авторизация телефонии не пройдена"); }

				try { RequestInternet_Post(); }
				catch { MessageBox.Show("Авторизация интернета не пройдена"); }

				Telefonia();
				Internet();

				label2.Content = "Последнее обновление:";
				label9.Content = DateTime.Now.ToShortTimeString() + " " + DateTime.Now.ToShortDateString();

				m_notifyIcon = new System.Windows.Forms.NotifyIcon();
				//m_notifyIcon.BalloonTipTitle = "Зоголовок сообщения";
				//m_notifyIcon.BalloonTipText = "Появляется когда мы помещаем иконку в трэй";

				m_notifyIcon.Text = "Телефония: \n" + sumCost + "\n\nИнтернет: \nIn: " + trafIn + "\nOut: " + trafOut;

				m_notifyIcon.Icon = new System.Drawing.Icon(@"tray.ico");
				m_notifyIcon.Click += new EventHandler(m_notifyIcon_Click);
			}
			catch
			{
				MessageBox.Show("Нет подключения к Internet...");
			}
		}
		#endregion

		#region Запрос авторизации телефонии
		public string RequestTelefonia_Post()
		{
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://tel.citytelecom.ru/?link=CL_CALLS");
			// Разрешаем авторедирект
			httpWebRequest.AllowAutoRedirect = true;
			// Создаем для запроса новый контейнер для хранения сессий
			httpWebRequest.CookieContainer = new CookieContainer();
			// Следующие строки итак понятны
			httpWebRequest.Method = "POST";
			httpWebRequest.ContentType = "application/x-www-form-urlencoded";
			// Переть тем как заполнять поля формы, текст конвертируем в байты
			byte[] ByteQuery = System.Text.Encoding.ASCII.GetBytes("username=" + tb_loginTel.Text + "&password=" + tb_passTel.Password);
			// Длинна запроса (обязательный параметр)
			httpWebRequest.ContentLength = ByteQuery.Length;
			// Открываем поток для записи
			Stream QueryStream = httpWebRequest.GetRequestStream();
			// Записываем в поток (это и есть POST запрос(заполнение форм))
			QueryStream.Write(ByteQuery, 0, ByteQuery.Length);
			// Закрываем поток
			QueryStream.Close();
			// Объект с ответом сервера
			HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
			// Присваиваем сессию
			httpWebResponse.Cookies = httpWebRequest.CookieContainer.GetCookies(httpWebRequest.RequestUri);
			if (httpWebResponse.Cookies != null)
			{
				// Добавляем сессию в наш контейнер для дальнейшего использования
				Cooks.Add(httpWebResponse.Cookies);
			}
			// Открываем поток для чтения
			Stream stream = httpWebResponse.GetResponseStream();
			// Читаем из потока
			StreamReader reader = new StreamReader(stream);
			// Возвращаем результат запроса
			return reader.ReadToEnd();
		}
		#endregion

		#region Получение данных телефонии
		public void Telefonia()
		{
			Date();
			var doc = new HtmlDocument();

			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://tel.citytelecom.ru/?link=CL_CALLS&sp=&pg=50&act=plain&begin_date=" + date1 + "&end_date=" + date2 + "&rep_type=S&duration=0&duration_max=0&charge_op_1=eq&charge=&charge_op=and&charge_op_2=eq&charge_max=&cdescription=&ccontype=&czone=&calling_number=&called_number=");
			httpWebRequest.AllowAutoRedirect = true;
			httpWebRequest.CookieContainer = new CookieContainer();
			if (Cooks != null)
			{
				//Добавляем к нашему запросу ранее сохраненную сессию
				httpWebRequest.CookieContainer.Add(Cooks);
			}
			HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
			httpWebResponse.Cookies = httpWebRequest.CookieContainer.GetCookies(httpWebRequest.RequestUri);
			try
			{
				using (var response = httpWebRequest.GetResponse())
				{
					using (var stream = response.GetResponseStream())
					{
						doc.Load(stream);

						numCalls_name = doc.DocumentNode.SelectSingleNode("html / body / center / table[4] / tr / td / table[2] / tr[1] / th[1]").InnerText;
						numCalls = doc.DocumentNode.SelectSingleNode("html / body / center / table[4] / tr / td / table[2] / tr[2] / td[1]").InnerText;

						sumCost_name = doc.DocumentNode.SelectSingleNode("html / body / center / table[4] / tr / td / table[2] / tr[1] / th[3]").InnerText.Remove(19);
						sumCost = doc.DocumentNode.SelectSingleNode("html / body / center / table[4] / tr / td / table[2] / tr[2] / td[3]").InnerText;

						sumDuration_name = doc.DocumentNode.SelectSingleNode("html / body / center / table[4] / tr / td / table[2] / tr[1] / th[2]").InnerText;
						sumDuration = doc.DocumentNode.SelectSingleNode("html / body / center / table[4] / tr / td / table[2] / tr[2] / td[2]").InnerText;

						textBox3.Text = numCalls_name + ": " + numCalls + "\n" + sumDuration_name + ": " + sumDuration + "\n" + sumCost_name + ": " + sumCost;
					}
				}
			}
			catch
			{
				textBox3.Text = "Не удалось получить данные";
			}
		}
		#endregion

		#region Запрос авторизации интернет
		public string RequestInternet_Post()
		{
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("https://billing.filanco.ru/index.php");
			// Разрешаем авторедирект
			httpWebRequest.AllowAutoRedirect = true;
			// Создаем для запроса новый контейнер для хранения сессий
			httpWebRequest.CookieContainer = new CookieContainer();
			// Следующие строки итак понятны
			httpWebRequest.Method = "POST";
			httpWebRequest.ContentType = "application/x-www-form-urlencoded";
			// Переть тем как заполнять поля формы, текст конвертируем в байты
			byte[] ByteQuery = System.Text.Encoding.ASCII.GetBytes("login=" + tb_loginInet.Text + "&password=" + tb_passInet.Password);
			// Длинна запроса (обязательный параметр)
			httpWebRequest.ContentLength = ByteQuery.Length;
			// Открываем поток для записи
			Stream QueryStream = httpWebRequest.GetRequestStream();
			// Записываем в поток (это и есть POST запрос(заполнение форм))
			QueryStream.Write(ByteQuery, 0, ByteQuery.Length);
			// Закрываем поток
			QueryStream.Close();
			// Объект с ответом сервера
			HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
			// Присваиваем сессию
			httpWebResponse.Cookies = httpWebRequest.CookieContainer.GetCookies(httpWebRequest.RequestUri);
			if (httpWebResponse.Cookies != null)
			{
				// Добавляем сессию в наш контейнер для дальнейшего использования
				Cooks.Add(httpWebResponse.Cookies);
			}
			// Открываем поток для чтения
			Stream stream = httpWebResponse.GetResponseStream();
			// Читаем из потока
			StreamReader reader = new StreamReader(stream);
			// Возвращаем результат запроса
			return reader.ReadToEnd();
		}
		#endregion

		#region Получение интернет данных
		private void Internet()
		{
			var doc = new HtmlDocument();
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("https://billing.filanco.ru/?devision=2");
			httpWebRequest.AllowAutoRedirect = true;
			httpWebRequest.CookieContainer = new CookieContainer();
			if (Cooks != null)
			{
				//Добавляем к нашему запросу ранее сохраненную сессию
				httpWebRequest.CookieContainer.Add(Cooks);
			}

			HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
			httpWebResponse.Cookies = httpWebRequest.CookieContainer.GetCookies(httpWebRequest.RequestUri);


			Stream responseStream = httpWebResponse.GetResponseStream();
			StreamReader responseStreamReader = new StreamReader(responseStream, Encoding.UTF8);


			try
			{
				doc.Load(responseStreamReader);

				trafIn_name = doc.DocumentNode.SelectSingleNode("html / body / table / tr[2] / td / table[3] / tr[2] / td[3]").InnerText;
				trafIn = doc.DocumentNode.SelectSingleNode("html / body / table / tr[2] / td / table[3] / tr[3] / td[3]").InnerText.Replace(".",",");

				trafOut_name = doc.DocumentNode.SelectSingleNode("html / body / table / tr[2] / td / table[3] / tr[2] / td[4]").InnerText;
				trafOut = doc.DocumentNode.SelectSingleNode("html / body / table / tr[2] / td / table[3] / tr[3] / td[4]").InnerText.Replace(".", ",");

				sumTraf_name = doc.DocumentNode.SelectSingleNode("html / body / table / tr[2] / td / table[3] / tr[2] / td[5]").InnerText;
				sumTraf = doc.DocumentNode.SelectSingleNode("html / body / table / tr[2] / td / table[3] / tr[3] / td[5]").InnerText.Replace(".", ",");

				textBox2.Text = trafIn_name + ": " + trafIn + "\n" + trafOut_name + ": " + trafOut + "\n" + sumTraf_name + ": " + sumTraf;

			}
			catch
			{
				textBox2.Text = "Не удалось получить данные";
			}
		}
		#endregion

		#region Задаем даты в ComboBox
		public void Date()
		{
			date1 = comboBox1.SelectedItem.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Year.ToString();
			date2 = comboBox2.SelectedItem.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Year.ToString();
		}
		#endregion

		# region Задаем значения для ComboBox

		private void ComboBox()
		{
			for (int i = 1; i <= DateTime.DaysInMonth(Convert.ToInt32(DateTime.Now.Year), Convert.ToInt32(DateTime.Now.Month)); i++)
			{
				comboBox1.Items.Add(i);
				comboBox2.Items.Add(i);
			}

			comboBox1.SelectedItem = DateTime.Now.Day;
			comboBox2.SelectedItem = DateTime.Now.Day;
		}
		#endregion

		#region Timer
		private void dispatcherTimer()
		{
			DispatcherTimer dispatcherTimer = new DispatcherTimer();
			dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
			dispatcherTimer.Interval = new TimeSpan(0, Convert.ToInt32(cmb_timer.SelectedItem.ToString()), 0);
			dispatcherTimer.Start();
		}

		private void dispatcherTimer_Tick(object sender, EventArgs e)
		{
			comboBox1.SelectedItem = DateTime.Now.Day;
			comboBox2.SelectedItem = DateTime.Now.Day;

			btn_refresh_Click(null, null);
		}

		private void ComboBox_Timer()
		{
			for (int i = 1; i <= 60; i++)
			{
				cmb_timer.Items.Add(i);
			}
			cmb_timer.SelectedItem = 5;
		}

		private void cmb_timer_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			dispatcherTimer();
		}
		#endregion

		#region Tray
		private WindowState m_storedWindowState = WindowState.Normal;
		void OnStateChanged(object sender, EventArgs args)
		{
			if (WindowState == WindowState.Minimized)
			{
				Hide();
			}
			else
				m_storedWindowState = WindowState;
		}
		void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs args)
		{
			CheckTrayIcon();
		}

		void m_notifyIcon_Click(object sender, EventArgs e)
		{
			Show();
			WindowState = m_storedWindowState;
		}
		void CheckTrayIcon()
		{
			ShowTrayIcon(!IsVisible);
		}

		void ShowTrayIcon(bool show)
		{
			if (m_notifyIcon != null)
				m_notifyIcon.Visible = show;
		}

		private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			m_notifyIcon.Dispose();
			m_notifyIcon = null;
		}
		#endregion

		#region Нажатие кнопки Отправить
		private void btn_sendMail_Click(object sender, RoutedEventArgs e)
		{
			SendMail users = new SendMail("email@addres.com.", "Filanko " + DateTime.Now.ToShortTimeString() + " " + DateTime.Now.ToShortDateString(), "Телефония:%0A" + numCalls_name + ": " + numCalls + "%0A" + sumDuration_name + ": " + sumDuration + "%0A" + sumCost_name + ": " + sumCost + "%0A%0AИнтернет:%0A" + trafIn_name + ": " + trafIn + "%0A" + trafOut_name + ": " + trafOut + "%0A" + sumTraf_name + ": " + sumTraf);
			Process.Start(users.SendLetter()); //в качестве параметра передавать надо не класс, а строку
		}
		#endregion

		#region Планировщик отправок писем
		private void Scheduler()
		{
			DispatcherTimer dt = new DispatcherTimer();
			dt.Tick += new EventHandler(Dt_Tick);
			dt.Interval = new TimeSpan(0, 1, 0);
			dt.Start();
		}

		private void Dt_Tick(object sender, EventArgs e)
		{
			double sumCostCheck = Convert.ToDouble(sumCost.Replace(" руб.", ""));
			double trafInCheck = Convert.ToDouble(trafIn);
			double trafOutCheck = Convert.ToDouble(trafOut);

			string schedulerTime = DateTime.Now.ToShortTimeString();
			if (schedulerTime == "8:00")
			{
				if (sumCostCheck >= 1000)
				{
					lbl_Tel.Foreground = Brushes.Red;
					lbl_Tel.Content = "Счет за телефонию превышает допустимый лимит 1000р.!";
					textBox3.Foreground = Brushes.Red;
				}

				if (trafInCheck >= 1500)
				{
					lbl_InetIn.Foreground = Brushes.Red;
					lbl_InetIn.Content = "Сумма входящего интернет трафика превышает \nдопустимый лимит в 1.5Gb!";
					textBox2.Foreground = Brushes.Red;
				}

				if (trafOutCheck >= 1500)
				{
					lbl_InetOut.Foreground = Brushes.Red;
					lbl_InetOut.Content = "Сумма исходящего интернет трафика превышает \nдопустимый лимит в 1.5Gb!";
					textBox2.Foreground = Brushes.Red;
				}

				else { }
				btn_sendMail_Click(null, null);
			}

			if (schedulerTime == "14:00")
			{
				if (sumCostCheck >= 3000)
				{
					lbl_Tel.Foreground = Brushes.Red;
					lbl_Tel.Content = "Счет за телефонию превышает допустимый лимит 3000р.!";
					textBox3.Foreground = Brushes.Red;
				}

				if (trafInCheck >= 4000)
				{
					lbl_InetIn.Foreground = Brushes.Red;
					MessageBox.Show("Сумма входящего интернет трафика превышает \nдопустимый лимит в 4Gb!");
					textBox2.Foreground = Brushes.Red;
				}

				if (trafOutCheck >= 4000)
				{
					lbl_InetOut.Foreground = Brushes.Red;
					lbl_InetOut.Content = "Сумма исходящего интернет трафика превышает \nдопустимый лимит в 4Gb!";
					textBox2.Foreground = Brushes.Red;
				}

				else { }
				btn_sendMail_Click(null, null);
			}

			if (schedulerTime == "19:50")
			{
				if (sumCostCheck >= 4000)
				{
					lbl_Tel.Foreground = Brushes.Red;
					lbl_Tel.Content = "Счет за телефонию превышает допустимый лимит 4000р.!";
					textBox3.Foreground = Brushes.Red;
				}

				if (trafInCheck >= 7000)
				{
					lbl_InetIn.Foreground = Brushes.Red;
					MessageBox.Show("Сумма входящего интернет трафика превышает \nдопустимый лимит в 7Gb!");
					textBox2.Foreground = Brushes.Red;
				}

				if (trafOutCheck >= 7000)
				{
					lbl_InetOut.Foreground = Brushes.Red;
					lbl_InetOut.Content = "Сумма исходящего интернет трафика превышает \nдопустимый лимит в 7Gb!";
					textBox2.Foreground = Brushes.Red;
				}

				else { }
				btn_sendMail_Click(null, null);
			}

		}
		#endregion
	}
}
