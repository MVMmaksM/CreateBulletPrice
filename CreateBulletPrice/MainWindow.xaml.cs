using CreateBulletPrice.Models;
using CreateBulletPrice.Services;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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

namespace CreateBulletPrice
{  
    public partial class MainWindow : Window
    {
        private List<PerechenModel> perechenKor;
        private List<PerechenModel> perechenPolny;
        private List<PriceCityModel> pricesByCity;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnLoadKor_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    var filePath = openFileDialog.FileName;
                    perechenKor = Bullet.GetPerechen(File.ReadExcel(filePath));
                    LblCountRowKor.Content += $" {perechenKor?.Count} записей";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }      
        private void BtnLoadPolny_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    var filePath = openFileDialog.FileName;
                    perechenPolny = Bullet.GetPerechen(File.ReadExcel(filePath));
                    LblCountRowPolny.Content += $" {perechenPolny?.Count} записей"; 
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        private void BtnLoadDataPrice_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    var filePath = openFileDialog.FileName;
                    pricesByCity = Bullet.GetPricesByCity(File.ReadExcel(filePath));
                    LblCountRowPrice.Content += $" {pricesByCity?.Count} записей";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }      
        private void CreatePolnyBullet(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.SpecialFolder.Desktop.ToString();
            saveFileDialog.FileName = $"Для бюллетеня полный от {DateTime.Now.Date.ToShortDateString()}";
            saveFileDialog.Filter = "*.xlsx|*.xlsx";
            try
            {
                if (saveFileDialog.ShowDialog() == true && perechenPolny is not null && pricesByCity is not null)
                {
                    var fileName = saveFileDialog.FileName;
                    var joinCollection = Bullet.JoinCollection(perechenPolny, pricesByCity);
                    File.SaveFile(fileName, File.GetBulletExcel(joinCollection));
                    
                    LblBulletPolnyCount.Content = $"Сформировано: {joinCollection.Count} записей";                    
                    MessageBox.Show($"Файл сохранен на: {fileName}", "Уведомление", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }      
        }

        private void CreateKorBullet_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.SpecialFolder.Desktop.ToString();
            saveFileDialog.FileName = $"Для бюллетеня короткий от {DateTime.Now.Date.ToShortDateString()}";
            saveFileDialog.Filter = "*.xlsx|*.xlsx";

            try
            {
                if (saveFileDialog.ShowDialog() == true && perechenKor is not null && pricesByCity is not null)
                {
                    var fileName = saveFileDialog.FileName;
                    var joinCollection = Bullet.JoinCollection(perechenKor, pricesByCity);
                    File.SaveFile(fileName, File.GetBulletExcel(joinCollection));

                    LblBulletKorCount.Content = $"Сформировано: {joinCollection.Count} записей";
                    MessageBox.Show($"Файл сохранен на: {fileName}", "Уведомление", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
    }
}
