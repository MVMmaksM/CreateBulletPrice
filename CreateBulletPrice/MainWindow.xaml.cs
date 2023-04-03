using CreateBulletPrice.Models;
using CreateBulletPrice.Services;
using Microsoft.EntityFrameworkCore;
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
        private List<PriceCityModel> pricesByCity;      
        public MainWindow()
        {
            InitializeComponent();

            try
            {
                using (var dbContext = new ApplicationContext())
                {
                    LblCountRowBdKor.Content = $"короткий перечень: {dbContext.Perechen_kor.Count()} записей";
                    LblCountRowBdPolny.Content = $"полный перечень: {dbContext.Perechen_polny.Count()} записей";
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
                List<PerechenModelPolny> perechenPolny;

                using (var dbContext = new ApplicationContext())
                {
                    perechenPolny = dbContext.Perechen_polny.ToList();
                }

                if (saveFileDialog.ShowDialog() == true && perechenPolny is not null && pricesByCity is not null)
                {
                    var fileName = saveFileDialog.FileName;
                    var joinCollection = Bullet.JoinCollection(perechenPolny.ToList<PerechenModel>(), pricesByCity);
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

            List<PerechenModelKor> perechenKor;

            using (var dbContext = new ApplicationContext())
            {
                perechenKor = dbContext.Perechen_kor.ToList();
            }

            try
            {
                if (saveFileDialog.ShowDialog() == true && perechenKor is not null && pricesByCity is not null)
                {
                    var fileName = saveFileDialog.FileName;
                    var joinCollection = Bullet.JoinCollection(perechenKor.ToList<PerechenModel>(), pricesByCity);
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

        private void MenuItem_Click_Load_Kor(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    var filePath = openFileDialog.FileName;
                    var perechenKor = Bullet.GetPerechenKor(File.ReadExcel(filePath));
                    int loadCountRow = default;

                    using (var dbContext = new ApplicationContext())
                    {
                        dbContext.Database.ExecuteSqlRaw("DELETE FROM [dbo].[Perechen_kor]");

                        foreach (var item in perechenKor)
                        {
                            dbContext.Perechen_kor.Add(item);
                            loadCountRow += dbContext.SaveChanges();

                            LblCountRowBdKor.Content = $"короткий перечень: {dbContext.Perechen_kor.Count()} записей";                         
                        }
                    }

                    MessageBox.Show($"Добавлено в БД в короткий перечень: {loadCountRow} записей", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void MenuItem_Click_Load_Polny(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    var filePath = openFileDialog.FileName;
                    var perechenPolny = Bullet.GetPerechenPolny(File.ReadExcel(filePath));
                    int loadCountRow = default;

                    using (var dbContext = new ApplicationContext())
                    {
                        dbContext.Database.ExecuteSqlRaw("DELETE FROM [dbo].[Perechen_polny]");

                        foreach (var item in perechenPolny)
                        {
                            dbContext.Perechen_polny.Add(item);
                            loadCountRow += dbContext.SaveChanges();

                            LblCountRowBdPolny.Content = $"полный перечень: {dbContext.Perechen_polny.Count()} записей";
                        }
                    }

                    MessageBox.Show($"Добавлено в БД в полный перечень: {loadCountRow} записей", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
    }
}
