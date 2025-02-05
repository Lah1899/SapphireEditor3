using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System;
using System.IO;
using Microsoft.Win32;

namespace SapphireEditor3
{
    public partial class MainWindow : Window
    {
        private string? currentFilePath;
        private bool isEdited;

        public MainWindow()
        {
            InitializeComponent();
            InitializeCommands();
            rtb.TextChanged += rtb_TextChanged;
            Closing += MainWindow_Closing;

            // コマンドライン引数があれば、最初の引数をファイルパスとして開く
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                OpenFile(args[1]);
            }
        }

        private void InitializeCommands()
        {
            // CommandBindings系の操作を追加するところ
            // GUI部品からのイベント駆動の実装時には忘れないようにね
            CommandBinding cb = new CommandBinding(ApplicationCommands.Paste, PasteResizedImage, PasteResizedImage_CanExecute);
            rtb.CommandBindings.Add(cb);
        }

        private void rtb_TextChanged(object sender, TextChangedEventArgs e)
        {
            // 編集されたら編集済みステータスを付加する
            if (!isEdited)
            {
                isEdited = true;
                UpdateTitle();
            }
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            // Ctrl + S で保存
            if (e.Key == Key.S && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                SaveFile(this, null);
                e.Handled = true;
            }

            // Ctrl + O でファイルを開く
            if (e.Key == Key.O && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                OpenFile(this, null);
                e.Handled = true;
            }

            // 以下はフォントスタイルの調節用。
            // 詳しい挙動については ApplyHeaderStyle メソッドを参照してくれ。

            // Ctrl + 1 で大見出しにする
            if ((e.Key == Key.D1 || e.Key == Key.NumPad1) && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                ApplyHeaderStyle(HeaderStyle.Large);
            }

            // Ctrl + 2 で中見出しにする
            if ((e.Key == Key.D2 || e.Key == Key.NumPad2) && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                ApplyHeaderStyle(HeaderStyle.Medium);
            }

            // Ctrl + 3 で小見出しにする
            if ((e.Key == Key.D3 || e.Key == Key.NumPad3) && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                ApplyHeaderStyle(HeaderStyle.Small);
            }

            // Ctrl + 0 で標準フォントに戻す
            if ((e.Key == Key.D0 || e.Key == Key.NumPad0) && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                ApplyHeaderStyle(HeaderStyle.Normal);
            }

            base.OnPreviewKeyDown(e);
        }

        private void PasteResizedImage_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            // ペースト操作を許可する
            // ラムダ式で書くことも考えたが、こっちのほうが処理が速そう
            e.CanExecute = true;
        }
        private void PasteResizedImage(object sender, ExecutedRoutedEventArgs e)
        {
            IDataObject dataObject = Clipboard.GetDataObject();
            if (dataObject != null && dataObject.GetDataPresent(DataFormats.Bitmap))
            {
                // クリップボードから画像を取得
                BitmapSource originalBitmap = Clipboard.GetImage();
                if (originalBitmap != null)
                {
                    // 画像面積が50000になるように伸縮させる
                    double targetArea = 50000;
                    double originalWidth = originalBitmap.PixelWidth;
                    double originalHeight = originalBitmap.PixelHeight;
                    double originalArea = originalWidth * originalHeight;
                    double scaleFactor = Math.Sqrt(targetArea / originalArea);
                    TransformedBitmap resizedBitmap = new TransformedBitmap(originalBitmap, new ScaleTransform(scaleFactor, scaleFactor));
                    originalBitmap = resizedBitmap;
                    // 調整した画像をクリップボードに設定
                    Clipboard.SetImage(originalBitmap);
                    // ペースト操作を実行
                    rtb.Paste();
                    // 元来のペースト処理が重複実行されないようにする
                    e.Handled = true;
                }
            }
            else
            {
                // ペースト対象が画像出なかった場合、通常のペースト処理を実行する。
                rtb.Paste();
                e.Handled = true;
            }
        }

        private void ApplyHeaderStyle(HeaderStyle headerStyle)
        {
            // 選択始点を含む行の行頭から、選択終点を含む行の行末まで、を処理対象とする
            TextPointer startOfSelection = rtb.Selection.Start;
            TextPointer endOfSelection = rtb.Selection.End;
            TextPointer startOfLine = startOfSelection.GetLineStartPosition(0);
            TextPointer endOfLine = endOfSelection.GetLineStartPosition(1);
            if (endOfLine == null)
            {
                // endOfLineが無効である場合、文章の末尾を代入してエラー回避
                endOfLine = rtb.Document.ContentEnd;
            }
            TextRange textRange = new TextRange(startOfLine, endOfLine);

            // 処理対象が空でないとき、処理対象のフォントを所定の形式に変更する
            if (!textRange.IsEmpty)
            {
                switch (headerStyle)
                {
                    case HeaderStyle.Large:
                        textRange.ApplyPropertyValue(TextElement.FontSizeProperty, 24.0);
                        textRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                        break;
                    case HeaderStyle.Medium:
                        textRange.ApplyPropertyValue(TextElement.FontSizeProperty, 20.0);
                        textRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                        break;
                    case HeaderStyle.Small:
                        textRange.ApplyPropertyValue(TextElement.FontSizeProperty, 16.0);
                        textRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                        break;
                    case HeaderStyle.Normal:
                        textRange.ApplyPropertyValue(TextElement.FontSizeProperty, 14.0);
                        textRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);
                        break;
                }
            }
        }

        private enum HeaderStyle
        {
            // 見出しの書式を管理するところ
            // Normalは標準のスタイルだよ
            Large,
            Medium,
            Small,
            Normal
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (isEdited)
            {
                // ウィンドウを閉じるとき編集中であればダイアログを表示
                MessageBoxResult result = MessageBox.Show("Save changes before closing?", "Confirmation", MessageBoxButton.YesNoCancel);

                if (result == MessageBoxResult.Yes)
                {
                    // Yes：ファイルを保存してアプリを終了する
                    if (!SaveFile(this, null))
                    {
                        // 保存に失敗したらアプリを終了しない
                        e.Cancel = true;
                    }
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    // Cancel：アプリが閉じないようにする
                    e.Cancel = true;
                }
            }
        }

        private void OpenFile(object sender, ExecutedRoutedEventArgs e)
        {
            // ファイルを開く
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "RTF Files|*.rtf|All Files|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                OpenFile(openFileDialog.FileName);
            }
        }

        private void OpenFile(string filePath)
        {
            try
            {
                TextRange textRange = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
                if (new System.IO.FileInfo(@filePath).Length == 0)
                {
                    // 0バイトのファイルは開かず、文字列なしとして扱う
                    textRange.Text = "";
                }
                else
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open))
                    {
                        textRange.Load(fs, DataFormats.Rtf);
                    }
                }
                currentFilePath = filePath;
                // 編集済みステータスの解除
                isEdited = false;
                // ウィンドウタイトルの更新
                UpdateTitle();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while opening the file: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool SaveFile(object sender, ExecutedRoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(currentFilePath))
            {
                return SaveAsFile(sender, e);
            }
            else
            {
                try
                {
                    using (FileStream fs = new FileStream(currentFilePath, FileMode.Create))
                    {
                        TextRange textRange = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
                        textRange.Save(fs, DataFormats.Rtf);
                    }
                    // 編集済みステータスを解除
                    isEdited = false;
                    UpdateTitle();
                    // 保存に成功したらTRUEを返す
                    return true;
                }
                catch (Exception ex)
                {
                    // 保存に失敗したらFALSEを返す
                    return false;
                }
            }
        }

        private bool SaveAsFile(object sender, ExecutedRoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "RTF Files|*.rtf|All Files|*.*";

            if (saveFileDialog.ShowDialog() == true)
            {
                String filePath = saveFileDialog.FileName;
                try
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Create))
                    {
                        TextRange textRange = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
                        textRange.Save(fs, DataFormats.Rtf);
                    }
                    // 編集済みステータスを解除
                    isEdited = false;
                    // 対象フルパスをcurrentPathに設定
                    currentFilePath = filePath;
                    UpdateTitle();
                    // 保存に成功したらTRUEを返す
                    return true;
                }
                catch (Exception ex)
                {
                    // 保存に失敗したらFALSEを返す
                    return false;
                }
            }
            // 保存に失敗したらFALSEを返す
            return false;
        }

        private void UpdateTitle()
        {
            string fileName = currentFilePath != null ? currentFilePath : "Untitled";
            string editedStatus = isEdited ? " (Edited)" : "";
            Title = $"{fileName}{editedStatus} - SapphireEditor 3.1.1";
        }
    }
}