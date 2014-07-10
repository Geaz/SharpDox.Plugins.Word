using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using System.IO;
using NotesFor.HtmlToOpenXml;
using System.Drawing;

namespace SharpDox.Plugins.Word.OpenXml.Elements
{
    internal class Image : BaseElement
    {
        private Size _imageSize;
        private Size _customSize;
        private MainDocumentPart _mainDocumentPart;

        public Image(string content, int? width = null, int? height = null) : base(content) 
        {
            GetImageSize(width, height);
        }

        public override void AppendTo(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            _mainDocumentPart = mainDocumentPart;

            var imagePart = CreateImagePart();
            openXmlNode.Append(CreateImageElement(mainDocumentPart.GetIdOfPart(imagePart), Path.GetFileName(_content)));
        }

        public override void InsertAfter(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            _mainDocumentPart = mainDocumentPart;

            var imagePart = CreateImagePart();
            openXmlNode.InsertAfterSelf(CreateImageElement(mainDocumentPart.GetIdOfPart(imagePart), Path.GetFileName(_content)));
        }

        private void GetImageSize(int? width, int? height)
        {
            using (var img = System.Drawing.Image.FromFile(_content))
            {
                _imageSize = new Size(img.Width, img.Height);
            }

            if(width == null && height == null)
            {
                _customSize = _imageSize;
            }
            else if(width != null && height != null)
            {
                _customSize = new Size(width.Value, height.Value);
            }
            else if(width != null)
            {
                _customSize = new Size(width.Value, ((_imageSize.Width / width.Value) * _imageSize.Height));
            }
            else
            {
                _customSize = new Size(((_imageSize.Height / height.Value) * _imageSize.Width), height.Value);
            }
        }

        private ImagePart CreateImagePart()
        {
            var imagePart = _mainDocumentPart.AddImagePart(ImagePartType.Png);
            using (var stream = new FileStream(_content, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            return imagePart;
        }

        private OpenXmlElement CreateImageElement(string relationshipId, string imageName)
        {
            var widthInEmus = new Unit(UnitMetric.Pixel, _customSize.Width).ValueInEmus;
            var heightInEmus = new Unit(UnitMetric.Pixel, _customSize.Height).ValueInEmus;
            var drawingId = GetNextDrawingId();
            var imageId = GetNextImageId();

            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = widthInEmus, Cy = heightInEmus },
                    new DW.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties() { Id = drawingId, Name = imageName, Description = string.Empty },
                    new DW.NonVisualGraphicFrameDrawingProperties
                    {
                        GraphicFrameLocks = new A.GraphicFrameLocks() { NoChangeAspect = true }
                    },
                    new A.Graphic(
                        new A.GraphicData(
                            new Pic.Picture(
                                new Pic.NonVisualPictureProperties
                                {
                                    NonVisualDrawingProperties = new Pic.NonVisualDrawingProperties() { Id = imageId, Name = imageName, Description = string.Empty },
                                    NonVisualPictureDrawingProperties = new Pic.NonVisualPictureDrawingProperties(
                                        new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
                                },
                                new Pic.BlipFill(
                                    new A.Blip() { Embed = relationshipId },
                                    new A.SourceRectangle(),
                                    new A.Stretch(
                                        new A.FillRectangle())),
                                new Pic.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = widthInEmus, Cy = heightInEmus }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList()
                                    ) { Preset = A.ShapeTypeValues.Rectangle }
                                ) { BlackWhiteMode = A.BlackWhiteModeValues.Auto })
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                ) { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U }
            );

            return element;
        }

        private UInt32Value GetNextDrawingId()
        {
            var drawingId = new UInt32Value((uint)1);
            foreach (var drawing in _mainDocumentPart.Document.Body.Descendants<Drawing>())
            {
                if (drawing.Inline != null && drawing.Inline.DocProperties.Id > drawingId)
                {
                    drawingId = drawing.Inline.DocProperties.Id;
                }
            }

            return drawingId > 1 ? ++drawingId : drawingId;
        }

        private UInt32Value GetNextImageId()
        {
            var imageId = new UInt32Value((uint)1);
            foreach (var drawing in _mainDocumentPart.Document.Body.Descendants<Drawing>())
            {
                var nvPr = drawing.Inline.Graphic.GraphicData.GetFirstChild<Pic.NonVisualPictureProperties>();
                if (nvPr != null && nvPr.NonVisualDrawingProperties.Id > imageId)
                {
                    imageId = nvPr.NonVisualDrawingProperties.Id;
                }
            }

            return imageId > 1 ? ++imageId : imageId;
        }
    }
}
