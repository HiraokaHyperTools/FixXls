using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FixXls {
    class Program {
        [STAThread]
        static void Main(string[] args) {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            OpenFileDialog ofdXls = new OpenFileDialog();
            ofdXls.Filter = "*.xls|*.xls";
            if (ofdXls.ShowDialog() != DialogResult.OK)
                return;

            SaveFileDialog sfdXls = new SaveFileDialog();
            sfdXls.DefaultExt = "xls";
            sfdXls.Filter = "*.xls|*.xls";
            if (sfdXls.ShowDialog() != DialogResult.OK)
                return;

            new Program().Recover(ofdXls.FileName, sfdXls.FileName);

            MessageBox.Show("修復しました。");
        }

        private void Recover(string fpSrc, string fpNew) {
            var si = new MemoryStream(File.ReadAllBytes(fpSrc));
            var wb = new HSSFWorkbook(si);
            var wb2 = new HSSFWorkbook();
            var df = wb.CreateDataFormat();
            var df2 = wb2.CreateDataFormat();
            SortedDictionary<int, ICellStyle> csMap = new SortedDictionary<int, ICellStyle>();
            SortedDictionary<int, IFont> csF = new SortedDictionary<int, IFont>();
            for (int z = 0; z < wb.NumberOfSheets; z++) {
                var sh = wb.GetSheetAt(z);
                var sh2 = new HSSFSheet(wb2);
                wb2.Add(sh2);
                wb2.SetSheetName(z, sh.SheetName);
                var y0 = sh.FirstRowNum;
                var y1 = sh.LastRowNum;
                SortedDictionary<int, int> useX = new SortedDictionary<int, int>();
                for (int t = 0; t < sh.NumMergedRegions; t++) {
                    sh2.AddMergedRegion(sh.GetMergedRegion(t));
                }
                for (int y = y0; y <= y1; y++) {
                    var row = sh.GetRow(y);
                    if (row == null) continue;
                    var row2 = sh2.CreateRow(y);
                    row2.Height = row.Height;
                    var x0 = row.FirstCellNum;
                    var x1 = row.LastCellNum;
                    if (x0 < 0) continue;
                    for (int x = x0; x <= x1; x++) {
                        useX[x] = 0;
                        var cell = row.GetCell(x);
                        if (cell == null) continue;
                        var cell2 = row2.CreateCell(x, cell.CellType);
                        switch (cell.CellType) {
                            case CellType.Boolean:
                                cell2.SetCellValue(cell.BooleanCellValue);
                                break;
                            case CellType.Numeric:
                                cell2.SetCellValue(cell.NumericCellValue);
                                break;
                            case CellType.String:
                                cell2.SetCellValue(cell.StringCellValue);
                                break;
                            case CellType.Formula:
                                cell2.SetCellFormula(cell.CellFormula);
                                //d["" + cell.CellFormula] = "";
                                break;
                        }
                        {
                            ICellStyle ns = null;
                            {
                                int iSrc = cell.CellStyle.Index;
                                if (!csMap.TryGetValue(iSrc, out ns)) {
                                    ns = wb2.CreateCellStyle();
                                    csMap[iSrc] = ns;
                                }
                            }
                            IFont font2 = null;
                            {
                                int iSrc = cell.CellStyle.FontIndex;
                                if (!csF.TryGetValue(iSrc, out font2)) {
                                    font2 = wb2.CreateFont();
                                    var font = cell.CellStyle.GetFont(wb);
                                    font2.Boldweight = font.Boldweight;
                                    font2.Charset = font.Charset;
                                    font2.Color = font.Color;
                                    font2.FontHeight = font.FontHeight;
                                    font2.FontName = font.FontName;
                                    font2.IsBold = font.IsBold;
                                    font2.IsItalic = font.IsItalic;
                                    font2.IsStrikeout = font.IsStrikeout;
                                    font2.TypeOffset = font.TypeOffset;
                                    font2.Underline = font.Underline;
                                    csF[iSrc] = font2;
                                }
                            }
                            ns.Alignment = cell.CellStyle.Alignment;
                            ns.BorderBottom = cell.CellStyle.BorderBottom;
                            ns.BorderDiagonal = cell.CellStyle.BorderDiagonal;
                            ns.BorderDiagonalColor = cell.CellStyle.BorderDiagonalColor;
                            ns.BorderDiagonalLineStyle = cell.CellStyle.BorderDiagonalLineStyle;
                            ns.BorderLeft = cell.CellStyle.BorderLeft;
                            ns.BorderRight = cell.CellStyle.BorderRight;
                            ns.BorderTop = cell.CellStyle.BorderTop;
                            ns.BottomBorderColor = cell.CellStyle.BottomBorderColor;
                            ns.DataFormat = df2.GetFormat(cell.CellStyle.GetDataFormatString()
                                ?? GetBuiltinFormat(cell.CellStyle.DataFormat)
                                );
                            ns.FillBackgroundColor = cell.CellStyle.FillBackgroundColor;
                            ns.FillForegroundColor = cell.CellStyle.FillForegroundColor;
                            ns.FillPattern = cell.CellStyle.FillPattern;
                            ns.SetFont(font2);
                            ns.Indention = cell.CellStyle.Indention;
                            ns.IsHidden = cell.CellStyle.IsHidden;
                            ns.IsLocked = cell.CellStyle.IsLocked;
                            ns.LeftBorderColor = cell.CellStyle.LeftBorderColor;
                            ns.RightBorderColor = cell.CellStyle.RightBorderColor;
                            ns.Rotation = cell.CellStyle.Rotation;
                            ns.ShrinkToFit = cell.CellStyle.ShrinkToFit;
                            ns.TopBorderColor = cell.CellStyle.TopBorderColor;
                            ns.VerticalAlignment = cell.CellStyle.VerticalAlignment;
                            ns.WrapText = cell.CellStyle.WrapText;

                            cell2.CellStyle = ns;
                        }
                    }
                }
                foreach (var x in useX.Keys)
                    sh2.SetColumnWidth(x, sh.GetColumnWidth(x));
            }
            using (var os = File.Create(fpNew)) {
                wb2.Write(os);
            }
        }

        private static string GetBuiltinFormat(short p) {
            try {
                return HSSFDataFormat.GetBuiltinFormat(p);
            }
            catch (ArgumentOutOfRangeException) {
                return "";
            }
        }
    }
}
