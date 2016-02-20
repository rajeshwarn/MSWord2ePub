using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace Word2HTML4ePub
{
    public abstract class Traitement_Images 
    {
        static ImageCodecInfo jgpEncoder = null;
        static ImageCodecInfo gifEncoder = null;
        static ImageCodecInfo pngEncoder = null;

        static EncoderParameters myEncoderParameters = null;


        private static void Init(int QualityLevel)
        {
            if (jgpEncoder == null)
                jgpEncoder = GetEncoder(ImageFormat.Jpeg);

            if (gifEncoder == null)
                gifEncoder = GetEncoder(ImageFormat.Gif);

            if (pngEncoder == null)
                pngEncoder = GetEncoder(ImageFormat.Png);
            
            if (myEncoderParameters == null)
                myEncoderParameters = new EncoderParameters(1);

            //// Create an Encoder object based on the GUID for the Quality parameter category.
            //System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;
            //EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 100);
            
            //EncoderParameter myEncoderParameter = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100);

            myEncoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, QualityLevel);
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        public static string ReduceBitmapDim(string SourceFile, int dimMax)
        {
            //Load bitmap
            Bitmap b = new Bitmap(SourceFile);

            float ratio = 1.0f;
            if (b.Height > b.Width)
            {
                if (b.Height > dimMax)
                    ratio = (float)dimMax / b.Height;
            }
            else
                if (b.Width > dimMax)
                    ratio = (float)dimMax / b.Width;

            //Reduce size
            if (ratio < 1.0f)
            {
                Bitmap bt = new Bitmap((int)(ratio * b.Width), (int)(ratio * b.Height), PixelFormat.Format24bppRgb);
                using (Graphics graphics = Graphics.FromImage(bt))
                {
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.DrawImage(b, 0, 0, (int)(ratio * b.Width), (int)(ratio * b.Height));
                }

                // Save the bitmap as a JPG file.
                string tempo = Path.GetTempFileName();
                bt.Save(tempo, ImageFormat.Jpeg);
                return tempo;
            }

            return SourceFile;
        }

        public static Size ReduceBitmapDim(string SourceFile, int dimMaxX, int dimMaxY)
        {
            //Load bitmap
            MemoryStream ms = new MemoryStream(File.ReadAllBytes(SourceFile));
            Bitmap b = new Bitmap(ms);

            if (b == null)
                return new Size(0,0);
            Size dim = b.Size;

            //Si l'image est plus petite, pas de modification
            if ((b.Width < dimMaxX) && (b.Height < dimMaxY))
                return new Size(b.Width, b.Height);

            //Calcul du facteur de réduction 
            float ratiow = (float)dimMaxX / b.Width;
            float ratioh = (float)dimMaxY / b.Height;

            float ratio = 1.0f;
            if (ratioh > ratiow)
                ratio = ratiow;
            else
                ratio = ratioh;
             
            //Reduce size
            if (ratio < 1.0f)
            {
                Bitmap bt = new Bitmap((int)(ratio * b.Width), (int)(ratio * b.Height), PixelFormat.Format24bppRgb);
                using (Graphics graphics = Graphics.FromImage(bt))
                {
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.DrawImage(b, 0, 0, (int)(ratio * b.Width), (int)(ratio * b.Height));
                }
                
                // Save the bitmap as a JPG file.
                Init(80);
                try
                { 
                    if (Path.GetExtension(SourceFile).ToLower().Contains("jp"))
                        bt.Save(SourceFile, jgpEncoder, myEncoderParameters);
                    else if (Path.GetExtension(SourceFile).ToLower().Contains("gif"))
                        bt.Save(SourceFile, gifEncoder, myEncoderParameters);
                    else if (Path.GetExtension(SourceFile).ToLower().Contains("png"))
                        bt.Save(SourceFile, pngEncoder, myEncoderParameters);
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Impossible de sauvegarder le fichier image!", "Erreur", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return dim;
                }
                return new Size(bt.Width, bt.Height); ;
            }

            return new Size(b.Width, b.Height);
        }


        /// <summary>
        /// Sauvegarde un fichier avec un taux de compression donné
        /// </summary>
        /// <param name="SourceFile">path vers le fichier</param>
        /// <param name="DestinationFile">path vers le nouveau fichier</param>
        /// <param name="encodingRatio">taux de compression initial</param>
        public static void CompressBitmap(string SourceFile, string DestinationFile, int encodingRatio)
        {
            Init(encodingRatio);

            //Load bitmap
            Bitmap b = new Bitmap(SourceFile);

            // Save the bitmap as a JPG file with dedicated quality level compression.
            b.Save(DestinationFile, jgpEncoder, myEncoderParameters);

        }

        /// <summary>
        /// Ajuste le niveau de compression pour que la taille d'une image soit inférieure à un seuil.
        /// </summary>
        /// <param name="SourceFile">path vers le fichier</param>
        /// <param name="DestinationFile">path vers le nouveau fichier</param>
        /// <param name="encodingRatio">taux de compression initial</param>
        /// <param name="MaxSize">Taille maxi</param>
        /// <returns>Taille du fichier</returns>
        public static long CompressBitmap(string SourceFile, string DestinationFile, int encodingRatio, long MaxSize)
        {
            //Load bitmap
            Bitmap b = new Bitmap(SourceFile);

            while (encodingRatio >= 0)
            {
                Init(encodingRatio);

                // Save the bitmap as a JPG file with dedicated quality level compression.
                using (MemoryStream ms = new MemoryStream())
                {
                    b.Save(ms, jgpEncoder, myEncoderParameters);
                    if (ms.Length < (MaxSize))
                    {
                        b.Save(DestinationFile, jgpEncoder, myEncoderParameters);
                        return ms.Length;
                    }
                    else
                    {
                        encodingRatio -= 10;
                    }
                }
            }
            return -1;
        }
    }
}
