using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Xml;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.SS.UserModel;
using System.Windows.Forms;
using ICell = NPOI.SS.UserModel.ICell;

namespace ProII_to_Excel
{
    class Conversion
    {
        private System.IO.Stream fileStream;
        private Dictionary<String, List<Point>> points;

        public Conversion(System.IO.Stream fileStream)
        {
            this.fileStream = fileStream;
            points = new Dictionary<string, List<Point>>();
            this.ReadFile();
        }

        private void ReadFile()
        {
            using (System.IO.StreamReader reader = new System.IO.StreamReader(this.fileStream))
            {
                String[] plt = reader.ReadToEnd().Split('\n');
                for (int i = 0; i < plt.Length; i++)
                {
                    String line = plt[i];
                    if (line.IndexOf("CURVE") != -1)
                    {
                        String title = line.Split('"')[1];
                        line = plt[++i];
                        line = stripAndTrim(plt[++i]);
                        List<Point> buffer = new List<Point>();
                        do
                        {
                            String[] split = line.Split(' ');
                            buffer.Add(new Point(XmlConvert.ToDouble(split[0]), XmlConvert.ToDouble(split[1])));
                            line = stripAndTrim(plt[++i]);
                        } while (line.IndexOf("MARKER") == -1);
                        points.Add(title, buffer);
                    }
                }

            }
            fileStream.Close();
        }

        public IWorkbook Convert()
        {
            if (points.Count > 0)
            {


                // Converts the data to Excel format
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet();
                IRow heading = sheet.CreateRow(0);
                heading.CreateCell(0).SetCellValue("Pressure/Temp");

                List<int> sizes = new List<int>();
                int size = 0;

                for (int i = 0; i < points.Count; i++)
                {
                    points.ElementAt(i).Value.Sort(delegate (Point a, Point b)
                    {
                        return a.TemperatureOrPressure.CompareTo(b.TemperatureOrPressure);
                    });
                    heading.CreateCell(i + 1).SetCellValue(points.ElementAt(i).Key);
                    sizes.Add(points.ElementAt(i).Value.Count);
                }
                size = sizes.Min();

                for (int i = 0; i < size; i++)
                {
                    IRow row = sheet.CreateRow(i + 1);

                    NPOI.SS.UserModel.ICell temp = row.CreateCell(0);
                    temp.SetCellType(CellType.Numeric);
                    temp.SetCellValue(points.ElementAt(0).Value.ElementAt(0).TemperatureOrPressure);

                    for (int j = 0; j < points.Count; j++)
                    {
                        ICell x = row.CreateCell(j + 1);
                        x.SetCellType(CellType.Numeric);
                        x.SetCellValue(points.ElementAt(j).Value.ElementAt(i).Fraction);
                    }

                }
                return wb;
            }
            else
            {
                MessageBox.Show("The selected file does not have any points", "Error");
            }
            return null;
        }

        private String stripAndTrim(String data)
        {
            return data.Trim();
        }

    }
}
