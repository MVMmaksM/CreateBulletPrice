using CreateBulletPrice.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace CreateBulletPrice.Services
{
    internal class Bullet
    {
        public static List<PerechenModelKor>? GetPerechenKor(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null)
            {
                return null;
            }

            var listPerechen = new List<PerechenModelKor>();

            for (int i = 1; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                var row = excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column];
                PerechenModelKor perechenModel = new PerechenModelKor();

                perechenModel.Ord = int.Parse(row[i, 1].Text);
                perechenModel.Kod = int.Parse(row[i, 2].Text);
                perechenModel.Name = row[i, 3].Text;

                listPerechen.Add(perechenModel);
            }

            return listPerechen;
        }

        public static List<PerechenModelPolny>? GetPerechenPolny(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null)
            {
                return null;
            }

            var listPerechen = new List<PerechenModelPolny>();

            for (int i = 1; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                var row = excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column];
                PerechenModelPolny perechenModel = new PerechenModelPolny();

                perechenModel.Ord = int.Parse(row[i, 1].Text);
                perechenModel.Kod = int.Parse(row[i, 2].Text);
                perechenModel.Name = row[i, 3].Text;

                listPerechen.Add(perechenModel);
            }

            return listPerechen;
        }

        public static List<PriceCityModel>? GetPricesByCity(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null)
            {
                return null;
            }

            var priceInCity = new List<PriceCityModel>();

            for (int i = 1; i < excelWorksheet.Dimension.End.Row; i++)
            {
                var row = excelWorksheet.Cells[i, 1, i, excelWorksheet.Dimension.End.Column];
                PriceCityModel priceCity = new PriceCityModel();

                priceCity.Name = row[i, 1].Text;
                priceCity.Kod = int.Parse(row[i, 2].Text);
                priceCity.Ufa = ConvertStringValue(row[i, 3].Text);
                priceCity.Ijevsk = ConvertStringValue(row[i, 4].Text);
                priceCity.Perm = ConvertStringValue(row[i, 5].Text);
                priceCity.Orenburg = ConvertStringValue(row[i, 6].Text);
                priceCity.Kurgan = ConvertStringValue(row[i, 7].Text);
                priceCity.Ekaterinburg = ConvertStringValue(row[i, 8].Text);
                priceCity.Tumen = ConvertStringValue(row[i, 9].Text);
                priceCity.Hanty = ConvertStringValue(row[i, 10].Text);
                priceCity.Salehard = ConvertStringValue(row[i, 11].Text);
                priceCity.Chelyabinsk = ConvertStringValue(row[i, 12].Text);

                priceInCity.Add(priceCity);
            }

            return priceInCity;
        }

        private static decimal? ConvertStringValue(string price)
        {
            if (decimal.TryParse(price, out decimal result))
            {
                return result;
            }
            else
            {
                return null;
            }
        }

        public static List<BulletModel> JoinCollection(List<PerechenModel> perechen, List<PriceCityModel> priceByCity)
        {
            return perechen.Join(priceByCity, per => per.Kod, price => price.Kod, (per, price) => new BulletModel()
            {
                Ord = per.Ord,
                Name = price.Name,
                Ufa = price.Ufa,
                Ijevsk = price.Ijevsk,
                Perm = price.Perm,
                Orenburg = price.Orenburg,
                Kurgan = price.Kurgan,
                Ekaterinburg = price.Ekaterinburg,
                Tumen = price.Tumen,
                Hanty = price.Hanty,
                Salehard = price.Salehard,
                Chelyabinsk = price.Chelyabinsk
            }).OrderBy(b => b.Ord).ToList();
        }
    }
}
