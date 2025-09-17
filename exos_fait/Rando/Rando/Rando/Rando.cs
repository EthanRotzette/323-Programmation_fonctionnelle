using Gpx;
using System;
using System.IO;

namespace Rando
{
    public partial class Rando : Form
    {
        public Rando()
        {
            InitializeComponent();
        }

        private void Rando_Form_Paint(object sender, PaintEventArgs e)
        {
            Pen myPen = new Pen(Color.Red);
            myPen.Width = 2;

            string filepath = "C:\\Users\\pk50gbi\\Documents\\GitHub\\323-Programmation_fonctionnelle\\exos_fait\\Rando\\gpx\\gemmikandersteg.gpx";
            var trackpoints = ReadFile(filepath);

            Point[] points = new Point[4] { new Point(30,50), new Point(50,10), new Point(80,50), new Point(111,400) };
            this.CreateGraphics().DrawLines(myPen, points);
        }
        private List<Trackpoint> ReadFile(string filePath)
        {
            List<Trackpoint> trackpoints = new List<Trackpoint>();
            var sr = new StreamReader(filePath);

            using (GpxReader reader = new GpxReader(sr.BaseStream))
            {
                while (reader.Read())
                {
                    switch (reader.ObjectType)
                    {
                        case GpxObjectType.Metadata:
                            //writer.WriteMetadata(reader.Metadata);
                            break;
                        case GpxObjectType.WayPoint:
                            //writer.WriteWayPoint(reader.WayPoint);
                            break;
                        case GpxObjectType.Route:
                            //writer.WriteRoute(reader.Route);
                            break;
                        case GpxObjectType.Track:
                            var data = reader.Track;
                            
                            var mappedData = data.ToGpxPoints().
                                Select(point => new Trackpoint() { Latitude = point.Latitude, Longitude = point.Longitude, Elevation = point.Elevation });

                            trackpoints.AddRange(mappedData);
                            break;
                    }
                }
            }

            return trackpoints;
        }
    }
    class Trackpoint
    {
        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public double? Elevation { get; set; }
    }
}
