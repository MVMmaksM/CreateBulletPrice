using CreateBulletPrice.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Shapes;

namespace CreateBulletPrice.Services
{
    internal class File
    {
        public static void SaveSetting(Setting settings) 
        {
            var settingString = JsonConvert.SerializeObject(settings);

            SaveFile($"{Environment.CurrentDirectory}\\Setting.json", Encoding.ASCII.GetBytes(settingString));
        }
        public static ExcelWorksheet ReadExcel(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(new FileInfo(path));

            var worksheet = package.Workbook.Worksheets[0];
            return worksheet;
        }

        public static byte[] GetBulletExcel(List<BulletModel> joinCollectionPrice)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage();

            var workSheetBullet = package.Workbook.Worksheets.Add("Бюллетень");
            workSheetBullet.Cells[1, 1, 2, 1].Merge = true;
            workSheetBullet.Cells[1, 2, 1, 5].Merge = true;
            workSheetBullet.Cells[1, 6, 1, 11].Merge = true;

            workSheetBullet.Cells[1, 1, 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[1, 1, 2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            workSheetBullet.Cells[1, 2, 1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[1, 6, 1, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetBullet.Cells[2, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            workSheetBullet.Cells[1, 1, 2, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheetBullet.Cells[1, 1, 2, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheetBullet.Cells[1, 1, 2, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheetBullet.Cells[1, 1, 2, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            workSheetBullet.Column(1).Width = 30;
            workSheetBullet.Column(5).Width = 10;
            workSheetBullet.Column(7).Width = 15;
            workSheetBullet.Column(9).Width = 16;
            workSheetBullet.Column(9).Width = 16;
            workSheetBullet.Column(11).Width = 10;

            workSheetBullet.Cells[3, 2, joinCollectionPrice.Count + 3, 11].Style.Numberformat.Format = "0.00";

            workSheetBullet.Cells[1, 1, 2, 1].Value = "Наименование";
            workSheetBullet.Cells[1, 2, 1, 5].Value = "Приволжский федеральный округ";
            workSheetBullet.Cells[1, 6, 1, 11].Value = "Уральский федеральный округ";
            workSheetBullet.Cells[2, 2].Value = "Уфа";
            workSheetBullet.Cells[2, 3].Value = "Ижевск";
            workSheetBullet.Cells[2, 4].Value = "Пермь";
            workSheetBullet.Cells[2, 5].Value = "Оренбург";
            workSheetBullet.Cells[2, 6].Value = "Курган";
            workSheetBullet.Cells[2, 7].Value = "Екатеринбург";
            workSheetBullet.Cells[2, 8].Value = "Тюмень";
            workSheetBullet.Cells[2, 9].Value = "Ханты-Мансийск";
            workSheetBullet.Cells[2, 10].Value = "Салехард";
            workSheetBullet.Cells[2, 11].Value = "Челябинск";

            for (int i = 0; i < joinCollectionPrice.Count; i++)
            {
                workSheetBullet.Cells[i + 3, 1].Value = joinCollectionPrice[i].Name;
                workSheetBullet.Cells[i + 3, 2].Value = joinCollectionPrice[i].Ufa;
                workSheetBullet.Cells[i + 3, 3].Value = joinCollectionPrice[i].Ijevsk;
                workSheetBullet.Cells[i + 3, 4].Value = joinCollectionPrice[i].Perm;
                workSheetBullet.Cells[i + 3, 5].Value = joinCollectionPrice[i].Orenburg;
                workSheetBullet.Cells[i + 3, 6].Value = joinCollectionPrice[i].Kurgan;
                workSheetBullet.Cells[i + 3, 7].Value = joinCollectionPrice[i].Ekaterinburg;
                workSheetBullet.Cells[i + 3, 8].Value = joinCollectionPrice[i].Tumen;
                workSheetBullet.Cells[i + 3, 9].Value = joinCollectionPrice[i].Hanty;
                workSheetBullet.Cells[i + 3, 10].Value = joinCollectionPrice[i].Salehard;
                workSheetBullet.Cells[i + 3, 11].Value = joinCollectionPrice[i].Chelyabinsk;
            }

            return package.GetAsByteArray();
        }

        public static void SaveFile(string pathFullNameFile, byte[] byteFile)
        {
            System.IO.File.WriteAllBytes(pathFullNameFile, byteFile);
        }
    }
}
