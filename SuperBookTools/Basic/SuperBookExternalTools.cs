#pragma warning disable CA2235 // Mark all non-serializable fields

using System;
using System.IO;

using IPA.Cores.Basic;
using IPA.Cores.Helper.Basic;
using static IPA.Cores.Globals.Basic;

using IPA.Cores.Codes;
using IPA.Cores.Helper.Codes;
using static IPA.Cores.Globals.Codes;

namespace SuperBookTools
{
    /// <summary>
    /// SuperBookTools external tools configuration
    /// Shared between CLI and GUI applications
    /// </summary>
    public static class SuperBookExternalTools
    {
        public static readonly ImageMagickUtil ImageMagick = new ImageMagickUtil(new ImageMagickOptions(
            Path.Combine(Env.AppRootDir, @"external_tools\image_tools\ImageMagick-portable-Q16-HDRI-x64\magick.exe"),
            Path.Combine(Env.AppRootDir, @"external_tools\image_tools\ImageMagick-portable-Q16-HDRI-x64\mogrify.exe"),
            Path.Combine(Env.AppRootDir, @"external_tools\image_tools\exiftool-13.30_64\exiftool.exe"),
            Path.Combine(Env.AppRootDir, @"external_tools\image_tools\QPDF\bin\qpdf.exe"),
            Path.Combine(Env.AppRootDir, @"external_tools\image_tools\pdfcpu\pdfcpu.exe")
        ));

        public static readonly FfMpegUtil FfMpeg = new FfMpegUtil(new FfMpegUtilOptions(
            Path.Combine(Env.AppRootDir, @"_dummy.exe"),
            Path.Combine(Env.AppRootDir, @"_dummy.exe")));

        public static readonly PdfYomitokuLib YomiToku = new PdfYomitokuLib(Path.Combine(Env.AppRootDir, @"external_tools\image_tools\yomitoku"));

        public static readonly AiUtilBasicSettings Settings = new AiUtilBasicSettings
        {
            AiTest_RealEsrgan_BaseDir = Path.Combine(Env.AppRootDir, @"external_tools\image_tools\RealEsrgan\RealEsrgan_Repo"),
            AiTest_TesseractOCR_Data_Dir = Path.Combine(Env.AppRootDir, @"external_tools\image_tools\TesseractOCR_Data"),
        };
        
        public static readonly AiTask Task = new AiTask(Settings, FfMpeg);

        public const string Post_OCR_Dir = "Post_OCR_Dir";
    }
}
