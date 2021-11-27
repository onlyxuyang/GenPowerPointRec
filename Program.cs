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

            using (var reader = new StreamReader(filepath))
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
                    float[] totalWeights = new float[count];
                    while (csv.Read())
                    {
                        slide = ppt.Slides[currentSlideIdx];
                        var record = Enumerable.ToList(csv.GetRecord<dynamic>());
                        if(record[0].Value == "")
                        {
                            currentx = startx;
                            currenty += height + gapy;
                            currentLineCount = 0;
                            continue;
                        }
                        if(((string)record[0].Value).ToLower() == "colors")
                        {
                            for (int i = 1; i < record.Count; i++)
                            {
                                //https://www.cnblogs.com/lv8218218/archive/2010/12/20/1911746.html
                                colors[i - 1] = System.Drawing.Color.FromName(record[i].Value);
                            }
                            continue;
                        }
                        float[] weights = new float[count];
                        for(int i=1;i<record.Count;i++)
                        {
                            weights[i-1] = float.Parse(record[i].Value);
                            totalWeights[i - 1] += weights[i - 1];
                        }
                        DrawRec(ref slide, record[0].Value, currentx, currenty, width, height, weights, colors);
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

                    DrawSamples(ref slide, csv.HeaderRecord.Skip(1).Take(csv.HeaderRecord.Length - 1).ToArray(), colors);
                    outputSummary(ref slide, csv.HeaderRecord.Skip(1).Take(csv.HeaderRecord.Length - 1).ToArray(), totalWeights);
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

        static void outputSummary(ref ISlide slide, string[] header, float[] weights)
        {
            string summary = "";
            for (int idx = 0; idx < header.Length; idx++)
            {
                summary += header[idx] + ":" + weights[idx].ToString()+"\n";
            }
            float width = 200;
            float height = 20;
            float startx = slide.Presentation.SlideSize.Size.Width - width - 20;
            float starty = 300;

            IAutoShape totalShape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(startx, starty, width, height));
            totalShape.Fill.FillType = FillFormatType.None;
            totalShape.TextFrame.AutofitType = TextAutofitType.Shape;
            totalShape.Line.FillType = FillFormatType.None;
            totalShape.TextFrame.Text = summary;
            TextRange textRange = totalShape.TextFrame.TextRange;
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = Color.Black;
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
