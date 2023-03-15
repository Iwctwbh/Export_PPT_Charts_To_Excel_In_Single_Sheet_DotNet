using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.Charts;

Console.WriteLine("For Grape");

// Licensing
Aspose.Cells.License licenseCells = new Aspose.Cells.License();

if (System.IO.File.Exists("Aspose.Total.lic"))
{
    using (Stream stream = System.IO.File.OpenRead("Aspose.Total.lic"))
    {
        licenseCells.SetLicense(stream);
    }
}

HashSet<int> unneededSlides = new HashSet<int>();

if (System.IO.File.Exists("unneededSlides.ini"))
{
    // 打开文件流
    using (StreamReader reader = new StreamReader("unneededSlides.ini"))
    {
        // 逐行读取文件内容
        string line;
        while ((line = reader.ReadLine()) != null)
        {
            if (line.Contains("-"))
            {
                for (int i = Convert.ToInt32(line.Split('-').FirstOrDefault()); i <= Convert.ToInt32(line.Split('-').LastOrDefault()); ++i)
                {
                    unneededSlides.Add(i);
                }
            }
            else
            {
                unneededSlides.Add(Convert.ToInt32(line));
            }
        }
    }
}

string folderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
string[] fileNames = Directory.GetFiles(folderPath).Select(Path.GetFileName).Where(w => w.EndsWith(".pptx")).ToArray();

foreach (string fileName in fileNames)
{
    // 加载 PPT 文件
    Console.WriteLine($"加载{fileName}文件。Loading...");

    using (var presentation = new Presentation(Path.Combine(fileName)))
    {

        // 创建一个新的 Excel 工作簿
        var workbook = new Workbook();
        // 获取第一个工作表
        var worksheet = workbook.Worksheets[0];

        int rowCount = 0;
        int rowTemp = 0;
        int columnTemp = 0;
        const int MAXROWCOUNT = 100;
        const int CHECKMAXROWCOUNT = 100;
        const int MAXCOLUMNCOUNT = 1000;

        // 循环遍历所有幻灯片
        foreach (var slide in presentation.Slides)
        {
            if (unneededSlides.Contains(slide.SlideNumber))
            {
                Console.WriteLine($"跳过第{slide.SlideNumber}页幻灯片。");
                continue;
            }
            // 循环遍历幻灯片上的所有形状
            foreach (var shape in slide.Shapes)
            {
                // 检查形状是否为图表
                if (shape is IChart)
                {
                    worksheet.Cells[rowCount, 0].Value = slide.SlideNumber;
                    Console.WriteLine($"遍历第{slide.SlideNumber}页幻灯片。");
                    // 获取图表对象
                    var chart = (IChart)shape;

                    // 获取图表数据
                    var chartData = chart.ChartData;

                    // 处理图表数据
                    var chartDataWorkbook = chartData.ChartDataWorkbook;

                    // 循环遍历 chartDataWorkSheet 中的所有项
                    for (rowTemp = 0; rowTemp < MAXROWCOUNT; ++rowTemp)
                    {
                        bool flag = false;
                        for (int i = 0; i < CHECKMAXROWCOUNT; ++i)
                        {
                            if (chartDataWorkbook.GetCell(0, rowTemp, i).Value != null)
                            {
                                flag = true;
                                break;
                            }
                        }
                        if (!flag)
                        {
                            break;
                        }

                        for (columnTemp = 1; columnTemp < MAXCOLUMNCOUNT; ++columnTemp)
                        {
                            // 获取单元格的值
                            var cellValue = chartDataWorkbook.GetCell(0, rowTemp, columnTemp - 1).Value;

                            // 将值赋给工作表的对应单元格
                            worksheet.Cells[rowTemp + rowCount, columnTemp].Value = cellValue;
                        }
                    }
                    rowCount += rowTemp + 2;
                }
            }
        }

        worksheet.AutoFitColumns();
        string fileSaveName = "Charts" + fileName.Substring(0, fileName.Length - 5) + ".xlsx";
        workbook.Save(fileSaveName);
        Console.WriteLine($"导出到{fileSaveName}!");
    }
}
Console.WriteLine("Complete!");
