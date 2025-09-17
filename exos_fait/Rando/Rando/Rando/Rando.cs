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
            Pen myPen = new Pen(Color.Red) { Width = 2 };

            string filepath = "C:\\Users\\pk50gbi\\Documents\\GitHub\\323-Programmation_fonctionnelle\\exos_fait\\Rando\\gpx\\gemmikandersteg.gpx";
            var trackpoints = ReadFile(filepath);

            if (trackpoints.Count == 0) return;

            // Récupérer les bornes GPS
            double minLat = trackpoints.Min(p => p.Latitude);
            double maxLat = trackpoints.Max(p => p.Latitude);
            double minLon = trackpoints.Min(p => p.Longitude);
            double maxLon = trackpoints.Max(p => p.Longitude);

            // Marge pour éviter que ça touche les bords
            int margin = 20;
            int width = this.ClientSize.Width - 2 * margin;
            int height = this.ClientSize.Height - 2 * margin;

            // Conversion GPS -> pixels
            Point[] points = trackpoints.Select(p =>
            {
                int x = margin + (int)((p.Longitude - minLon) / (maxLon - minLon) * width);

                int y = margin + (int)((maxLat - p.Latitude) / (maxLat - minLat) * height); // inversion Y

                return new Point(x, y);
            }).ToArray();

            e.Graphics.DrawLines(myPen, points);
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
                                Select(point => new Trackpoint(point.Latitude,point.Longitude,point.Elevation));

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
        public Trackpoint(double latitude, double longitude, double? elevation)
        {
            Latitude = latitude;
            Longitude = longitude;
            Elevation = elevation;
        }

        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public double? Elevation { get; set; }
    }
}
