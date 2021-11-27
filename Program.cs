using CsvHelper;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenPowerPointRec
{
    class Program
    {
        static void Main(string[] args)
        {
            string filepath = args[0];

            Presentation ppt = new Presentation();
            ppt.SlideSize.Type = SlideSizeType.Screen16x9;

            FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using (var reader = new StreamReader(fs))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    float startx = 50;
                    float starty = 50;
                    float width = 100;
                    float height = 30;
                    float gapx = 5;
                    float gapy = 5;
                    int linecount = 5;

                    float currentx = startx;
                    float currenty = starty;
                    int currentLineCount = 0;
                    int currentSlideIdx = 0;
 
                    csv.Read();
                    csv.ReadHeader();
                    int count = csv.HeaderRecord.Length - 1;

                    Color[] colors = new Color[count];
                    ISlide slide = ppt.Slides[currentSlideIdx];
                    Dictionary<string, Dictionary<string, float>> groupWeights = new Dictionary<string, Dictionary<string, float>>();
                    while (csv.Read())
                    {
                        slide = ppt.Slides[currentSlideIdx];
                        var record = Enumerable.ToList(csv.GetRecord<dynamic>());
                        if(record[0].Value == "")
                        {
                            if (currentLineCount == 0)
                                continue;

                            currentx = startx;
                            currenty += height + gapy;
                            currentLineCount = 0;
                            continue;
                        }
                        if(((string)record[0].Value).ToLower() == "colors")
                        {
                            for (int i = 1; i < record.Count; i++)
                            {
                                if(((string)record[i].Value).StartsWith("#"))
                                {
                                    int rgb = (int)(int.Parse(((string)record[i].Value).Remove(0,1), System.Globalization.NumberStyles.HexNumber) + 0xFF000000);
                                    colors[i - 1] = System.Drawing.Color.FromArgb(rgb);
                                }
                                else if(((string)record[i].Value).Contains(","))
                                {
                                    var rgb = ((string)record[i].Value).Split(',');
                                    colors[i - 1] = System.Drawing.Color.FromArgb(int.Parse(rgb[0]), int.Parse(rgb[1]), int.Parse(rgb[2]));
                                }
                                else
                                {
                                    //https://www.cnblogs.com/lv8218218/archive/2010/12/20/1911746.html
                                    colors[i - 1] = System.Drawing.Color.FromName(record[i].Value);
                                }
                            }
                            continue;
                        }
                        if (((string)record[0].Value).ToLower() == "size")
                        {
                            width = int.Parse(record[1].Value);
                            height = int.Parse(record[2].Value);
                            gapx = int.Parse(record[3].Value);
                            gapy = int.Parse(record[4].Value);
                            linecount = int.Parse(record[5].Value);
                            continue;
                        }

                        float[] weights = new float[count];
                        for(int i=2;i<record.Count;i++)
                        {
                            weights[i-2] = float.Parse(record[i].Value);
                            string group = record[0].Value;
                            string type = csv.HeaderRecord[i];
                            if(!groupWeights.ContainsKey(group))
                            {
                                groupWeights[group] = new Dictionary<string, float>();
                            }
                            if(!groupWeights[group].ContainsKey(type))
                            {
                                groupWeights[group][type] = 0;
                            }
                            groupWeights[group][type] += weights[i - 2];
                        }

                        DrawRec(ref slide, record[1].Value, currentx, currenty, width, height, weights, colors);
                        currentLineCount += 1;
                        if(currentLineCount >= linecount)
                        {
                            currentLineCount = 0;
                            currenty += height + gapy;
                            currentx = startx;
                        }
                        else
                        {
                            currentx += width + gapx;
                        }
                    }

                    DrawSamples(ref slide, csv.HeaderRecord.Skip(2).Take(csv.HeaderRecord.Length - 2).ToArray(), colors);
                    outputSummary(ref slide, groupWeights);
                }
            }

            string outputpath = Path.GetFileNameWithoutExtension(filepath) + ".pptx";
            ppt.SaveToFile(outputpath, FileFormat.Pptx2010);
        }

        static List<Color> GetAllColors()
        {
            List<Color> colors = new List<Color>();
            foreach (var item in typeof(Color).GetMembers())
            {
                if (item.MemberType == System.Reflection.MemberTypes.Property && System.Drawing.Color.FromName(item.Name).IsKnownColor == true)//只取属性且为属性中的已知Color，剔除byte属性以及一些布尔属性等（A B G R IsKnownColor Name等）
                {

                    colors.Add(System.Drawing.Color.FromName(item.Name));
                }
            }
            return colors;
        }

        static void outputSummary(ref ISlide slide, Dictionary<string, Dictionary<string, float>> groupWeights)
        {
            int rowNum = groupWeights.Values.First().Count + 1;
            int colNum = groupWeights.Count + 1; 
            double cellWidth = 80.0;
            double cellHeight = 20.0;
            double[] widths = Enumerable.Repeat(cellWidth, colNum).ToArray();
            double[] heights = Enumerable.Repeat(cellHeight, rowNum).ToArray();
            widths[0] = 160.0;
            double tableWidth = cellWidth * colNum;
            float startx = slide.Presentation.SlideSize.Size.Width - (float)tableWidth - 20;
            float starty = 300;

            ITable table = slide.Shapes.AppendTable(startx, starty, widths, heights);
            table.StylePreset = TableStylePreset.MediumStyle3Accent1;

            table[0, 0].TextFrame.Text = "";

            int colIdx = 0;
            int rowIdx = 0;
            foreach(var group in groupWeights)
            {
                colIdx++;
                table[colIdx, rowIdx].TextFrame.Text = group.Key;
            }
            colIdx = 0;
            foreach (var group in groupWeights)
            {
                colIdx++;
                rowIdx = 0;
                foreach (var type in group.Value)
                {
                    rowIdx++;
                    table[0, rowIdx].TextFrame.Text = type.Key;
                    table[colIdx, rowIdx].TextFrame.Text = type.Value.ToString();
                }
            }
        }
        static void DrawSamples(ref ISlide slide, string[] header, Color[] colors)
        {
            float width = 100;
            float height = 15;
            float startx = slide.Presentation.SlideSize.Size.Width - width - 20;
            float starty = 50;
            float gapy = 5;

            float currenty = starty;
            for (int idx = 0; idx < header.Length; idx++)
            {
                IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(startx, currenty, width, height));
                shape.Fill.FillType = FillFormatType.Solid;
                shape.Fill.SolidColor.Color = colors[idx];
                shape.Line.FillType = FillFormatType.None;
                shape.TextFrame.AutofitType = TextAutofitType.Normal;
                shape.TextFrame.Text = header[idx];
                currenty += height + gapy;
            }
        }
        static void DrawRec(ref ISlide slide, string title, float x, float y, float width, float height, float[] weights, Color[] colors)
        {
            float weightSum = 0;
            foreach (var weight in weights)
                weightSum += weight;
            float currentX = x;
            for(int idx = 0; idx<weights.Length; idx++)
            {
                if (weights[idx] == 0)
                    continue;
                float recWidth = (weights[idx] / (weightSum * 1.0f)) * width;
                IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(currentX, y, recWidth, height));
                shape.Fill.FillType = FillFormatType.Solid;
                shape.Fill.SolidColor.Color = colors[idx];
                shape.Line.FillType = FillFormatType.None;
                currentX += recWidth;
            }
            IAutoShape titleShape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(x, y, width, height));
            titleShape.Fill.FillType = FillFormatType.None;
            titleShape.TextFrame.AutofitType = TextAutofitType.Normal;
            titleShape.Line.FillType = FillFormatType.None;
            titleShape.TextFrame.Text = title;
        }
    }
}
