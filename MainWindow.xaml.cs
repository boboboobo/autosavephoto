using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs; // ★ パッケージインストール後に利用可能
using System; // IntPtr, Exceptionなど
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.Design;
using System.Windows.Input;
using System.Windows.Interop; // HwndSource関連
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace autosavephoto
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 


    public partial class MainWindow : Window
    {


        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        private const int WM_CLIPBOARDUPDATE = 0x031D;// OSからクリップボード更新メッセージが送られてきたときの定数

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // 1. WPFウィンドウハンドル（IntPtr）を取得
            WindowInteropHelper helper = new WindowInteropHelper(this);
            IntPtr hWnd = helper.Handle;

            // 2. Windowsメッセージをフックする処理を登録
            HwndSource source = HwndSource.FromHwnd(hWnd);
            source?.AddHook(WndProc); // sourceがnullでないことを確認

            // 3. クリップボードの変更を監視するようにOSに登録
            AddClipboardFormatListener(hWnd);
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            // クリップボードの監視を停止
            WindowInteropHelper helper = new WindowInteropHelper(this);
            RemoveClipboardFormatListener(helper.Handle);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // クリップボード更新メッセージが届いたかチェック
            if (msg == WM_CLIPBOARDUPDATE)
            {
                // ここにクリップボードから画像をロードする処理を呼び出す
                CheckAndLoadImageFromClipboard();
                handled = true;
            }
            return IntPtr.Zero;
        }

        private void CheckAndLoadImageFromClipboard()
        {
            // 画像をロードして表示する処理
            if (System.Windows.Clipboard.ContainsImage())
            {
                // ... (画像を取得し、DisplayImageControl.Sourceに設定する処理) ...

                try
                {
                    // 1. クリップボードから画像データ（BitmapSource）を取得
                    // BitmapSourceは、WPFのImageコントロールがSourceとして受け入れられるデータ形式
                    BitmapSource imageSource = System.Windows.Clipboard.GetImage();

                    // 2. 取得した画像をImageコントロールに設定する
                    // XAMLで <Image Name="ClipboardImage" ... /> のように定義されている前提
                    ClipboardImage.Source = imageSource;
                }
                catch (Exception ex)
                {
                    // 画像データが破損しているなどの例外を処理
                    MessageBox.Show($"クリップボードからの画像読み込み中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void OpenPath_Click(object sender, RoutedEventArgs e)
        {
            // 1. ダイアログオブジェクトを新規作成
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();

            // 2. ★これが重要★ ファイルではなく、フォルダを選択するように設定
            dialog.IsFolderPicker = true;

            // 3. ダイアログを表示
            CommonFileDialogResult result = dialog.ShowDialog();

            // 4. 結果をチェック
            if (result == CommonFileDialogResult.Ok)
            {
                // ユーザーがフォルダを選択し、「OK」を押した場合
                string selectedFolderPath = dialog.FileName; // FileNameプロパティにフォルダパスが入る

                // 選択されたパスを画面のTextBoxに表示
                FilePath.Text = selectedFolderPath;
            }
            else
            {
                // ユーザーがキャンセルした場合
                // 特に何もしないか、またはTextBoxの内容をクリアします
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            BitmapSource image = System.Windows.Clipboard.GetImage();
            BitmapEncoder encoder = null;

            if (ExtensionCombo.Text == ".png")
            {
                encoder = new PngBitmapEncoder();
            }
            else if (ExtensionCombo.Text == ".jpg")
            {
                // JPEGエンコーダを使用（画質はデフォルトで設定されます）
                JpegBitmapEncoder jpegEncoder = new JpegBitmapEncoder();

                // 必要に応じてQualityレベルを設定できます (0～100, デフォルトは90)
                // jpegEncoder.QualityLevel = 90; 

                encoder = jpegEncoder;
            }
            else
            {
                return;
            }

            // エンコーダに画像フレームを追加
            encoder.Frames.Add(BitmapFrame.Create(image));

            string fullFileName = FilePath.Text + @"\" + StartNumber.Text + BaseName.Text + ExtensionCombo.Text;

            using (FileStream fileStream = new FileStream(fullFileName, FileMode.Create))
            {
                encoder.Save(fileStream);
            }

            StartNumber.Text = (int.Parse(StartNumber.Text) + 1).ToString();

            LogText.Text += $"{fullFileName}に保存しました。\n";
            LogText.ScrollToEnd();
        }
    }

    


}