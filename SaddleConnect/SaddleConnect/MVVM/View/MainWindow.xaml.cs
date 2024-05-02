using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
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
using Autodesk.Revit.UI;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using TIGUtility;

namespace SaddleConnect
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region UI
        bool _isCancel = false;
        bool _isStoped = false;
        public async void EscFunction(Window window)
        {
            bool isValid = await loop();
            if (isValid)
            {
                _isCancel = false;
                _isStoped = false;
                //window.Close();
            }

        }
        public async Task<bool> loop()
        {
            await Task.Run(() =>
            {
                do
                {
                    try
                    {

                        _isStoped = _isCancel ? true : _isStoped;
                    }
                    catch (Exception)
                    {
                    }
                }
                while (!_isStoped);

            });
            return _isStoped;
        }
        private void InitializeMaterialDesign()
        {
            var card = new Card();
            var hue = new Hue("Dummy", Colors.Black, Colors.White);

        }
        private void InitializeWindowProperty()
        {
           
            //this.Title = Util.ApplicationWindowTitle;
            this.MinHeight = Util.ApplicationWindowHeight;
            this.Height = Util.ApplicationWindowHeight;
            this.Topmost = Util.IsApplicationWindowTopMost;
            this.MinWidth = Util.IsApplicationWindowAlowToReSize ? Util.ApplicationWindowWidth : 100;
            this.Width = Util.ApplicationWindowWidth;
            this.ResizeMode = Util.IsApplicationWindowAlowToReSize ? System.Windows.ResizeMode.CanResize : System.Windows.ResizeMode.NoResize;
            this.WindowStyle = WindowStyle.None;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            HeaderPanel.Instance = this;
        }
        #endregion

        public static MainWindow Instance;
        public UIApplication _UIApp = null;
        readonly List<ExternalEvent> _externalEvents = new List<ExternalEvent>();
        public MainWindow(CustomUIApplication application)
        {
            InitializeWindowProperty();
            InitializeMaterialDesign();
            InitializeComponent();
            InitializeHandlers();
            Instance = this;
            HeaderPanel.Instance = this;
            string path = System.IO.Path.GetDirectoryName(
               System.Reflection.Assembly.GetExecutingAssembly().Location);
            DirectoryInfo di = new DirectoryInfo(path);
            path = new DirectoryInfo(new DirectoryInfo(di.Parent.FullName).Parent.FullName).FullName + "\\HelpDocs\\HTML\\snvaddins";
            path = System.IO.Path.Combine(path, "ConPak");
            path = System.IO.Path.Combine(path, "SaddleConnect.html");
            HeaderPanel.DocumentPath = path;
            FooterPanel.Version = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
            //start EscFunction
            _isCancel = false;
            HotkeysManager.SetupSystemHook();
            HotkeysManager.AddHotkey(new GlobalHotkey(ModifierKeys.None, Key.Escape, () => { _isCancel = true; }));
            EscFunction(this);
            // end EscFunction
            UserControl userControl = new ParentUserControl(_externalEvents, application, this);
            Container.Children.Add(userControl);
        }
        private void InitializeHandlers()
        {
           
            _externalEvents.Add(ExternalEvent.Create(new saddlectHadler()));
          
        }
    }
}
