using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OC_Porject_20191112
{
    class DigitalGamma
    {
        public DataTable dt;
        public bool RGB;
        public UInt16 datawidth;
        public UInt16 addresswidth;
        public double m_x;
        public double m_y;
        public double tarlum;
        public double gamma;
        public UInt16 maxlum_limit;
        public byte[] flash = new byte[16400];
        public double r_lumfinal;
        public double g_lumfinal;
        public double b_lumfinal;
        public double w_lumfinal;
        public double α;
        public double β;
        public double γ;


        double[] R_LUM = new double[256];
        double[] G_LUM = new double[256];
        double[] B_LUM = new double[256];
        double[] W_LUM = new double[256];
        double r_x;
        double r_y;
        double g_x;
        double g_y;
        double b_x;
        double b_y;
        double w_x;
        double w_y;
        double Lr;
        double Lg;
        double Lb;
        double Lw;
        

        public void DataCache(DataTable dt)
        {
            for (int i = 0; i < 256; i++)
            {
                R_LUM[i] = Convert.ToDouble(dt.Rows[2 + i][6]);
                G_LUM[i] = Convert.ToDouble(dt.Rows[258 + i][6]);
                B_LUM[i] = Convert.ToDouble(dt.Rows[514 + i][6]);
                if (RGB)
                {
                    W_LUM[i] = 0;
                }
                else
                {
                    W_LUM[i] = Convert.ToDouble(dt.Rows[770 + i][6]);
                }
            }
            r_x = Convert.ToDouble(dt.Rows[257][4]);
            r_y = Convert.ToDouble(dt.Rows[257][5]);
            g_x = Convert.ToDouble(dt.Rows[513][4]);
            g_y = Convert.ToDouble(dt.Rows[513][5]);
            b_x = Convert.ToDouble(dt.Rows[769][4]);
            b_y = Convert.ToDouble(dt.Rows[769][5]);
            if (RGB)
            {
                w_x = 1;
                w_y = 1;
            }
            else
            {
                w_x = Convert.ToDouble(dt.Rows[1025][4]);
                w_y = Convert.ToDouble(dt.Rows[1025][5]);
            }
        }

        public void MaxLumLimit()
        {
            LrLgLbLwCal();
            UInt16 mll1 = (UInt16)(R_LUM[255] / Lr);
            UInt16 mll2 = (UInt16)(G_LUM[255] / Lg);
            UInt16 mll3 = (UInt16)(B_LUM[255] / Lb);
            UInt16 mll4;

            if (RGB)
            {
                mll4 = 0;
                if ((mll1 <= mll2) && (mll1 <= mll3))
                {
                    maxlum_limit = mll1;
                }
                else if ((mll2 <= mll1) && (mll2 <= mll3))
                {
                    maxlum_limit = mll2;
                }
                else
                {
                    maxlum_limit = mll3;
                }
            }
            else
            {
                mll4 = (UInt16)(W_LUM[255] / Lw);
                if ((mll1 <= mll2) && (mll1 <= mll3) && (mll1 <= mll4))
                {
                    maxlum_limit = mll1;
                }
                else if ((mll2 <= mll1) && (mll2 <= mll3) && (mll2 <= mll4))
                {
                    maxlum_limit = mll2;
                }
                else if ((mll3 <= mll1) && (mll3 <= mll2) && (mll3 <= mll4))
                {
                    maxlum_limit = mll3;
                }
                else
                {
                    maxlum_limit = mll4;
                }
            }
        }

        public void LumFinalCal()
        {
            LrLgLbLwCal();
            r_lumfinal = tarlum * Lr;
            g_lumfinal = tarlum * Lg;
            b_lumfinal = tarlum * Lb;
            w_lumfinal = tarlum * Lw;
        }

        public void LrLgLbLwCal()
        {
            double a1 = r_x / r_y;
            double b1 = g_x / g_y;
            double c1 = b_x / b_y;
            double d1 = w_x / w_y;
            double e1 = m_x / m_y;
            //double e1 = Convert.ToDouble(nUD_mx.Text) / Convert.ToDouble(nUD_my.Text);

            double a2 = (1 - r_x - r_y) / r_y;
            double b2 = (1 - g_x - g_y) / g_y;
            double c2 = (1 - b_x - b_y) / b_y;
            double d2 = (1 - w_x - w_y) / w_y;
            double e2 = (1 - m_x - m_y) / m_y;
            //double e2 = (1 - Convert.ToDouble(nUD_mx.Text) - Convert.ToDouble(nUD_my.Text)) / Convert.ToDouble(nUD_my.Text);

            double L1_1 = (e1 * c2 + c1 * d2 + d1 * e2 - e1 * d2 - c1 * e2 - d1 * c2) / (b1 * c2 + c1 * d2 + d1 * b2 - b1 * d2 - c1 * b2 - d1 * c2);
            double L1_2 = (b1 * e2 + e1 * d2 + d1 * b2 - b1 * d2 - e1 * b2 - d1 * e2) / (b1 * c2 + c1 * d2 + d1 * b2 - b1 * d2 - c1 * b2 - d1 * c2);
            double L1_3 = (b1 * c2 + c1 * e2 + e1 * b2 - b1 * e2 - c1 * b2 - e1 * c2) / (b1 * c2 + c1 * d2 + d1 * b2 - b1 * d2 - c1 * b2 - d1 * c2);
            double L2_1 = (e1 * c2 + c1 * d2 + d1 * e2 - e1 * d2 - c1 * e2 - d1 * c2) / (a1 * c2 + c1 * d2 + d1 * a2 - a1 * d2 - c1 * a2 - d1 * c2);
            double L2_2 = (a1 * e2 + e1 * d2 + d1 * a2 - a1 * d2 - e1 * a2 - d1 * e2) / (a1 * c2 + c1 * d2 + d1 * a2 - a1 * d2 - c1 * a2 - d1 * c2);
            double L2_3 = (a1 * c2 + c1 * e2 + e1 * a2 - a1 * e2 - c1 * a2 - e1 * c2) / (a1 * c2 + c1 * d2 + d1 * a2 - a1 * d2 - c1 * a2 - d1 * c2);
            double L3_1 = (e1 * b2 + b1 * d2 + d1 * e2 - e1 * d2 - b1 * e2 - d1 * b2) / (a1 * b2 + b1 * d2 + d1 * a2 - a1 * d2 - b1 * a2 - d1 * b2);
            double L3_2 = (a1 * e2 + e1 * d2 + d1 * a2 - a1 * d2 - e1 * a2 - d1 * e2) / (a1 * b2 + b1 * d2 + d1 * a2 - a1 * d2 - b1 * a2 - d1 * b2);
            double L3_3 = (a1 * b2 + b1 * e2 + e1 * a2 - a1 * e2 - b1 * a2 - e1 * b2) / (a1 * b2 + b1 * d2 + d1 * a2 - a1 * d2 - b1 * a2 - d1 * b2);

            Lr = (e1 * b2 + b1 * c2 + c1 * e2 - e1 * c2 - b1 * e2 - c1 * b2) / (a1 * b2 + b1 * c2 + c1 * a2 - a1 * c2 - b1 * a2 - c1 * b2);
            Lg = (a1 * e2 + e1 * c2 + c1 * a2 - a1 * c2 - e1 * a2 - c1 * e2) / (a1 * b2 + b1 * c2 + c1 * a2 - a1 * c2 - b1 * a2 - c1 * b2);
            Lb = (a1 * b2 + b1 * e2 + e1 * a2 - a1 * e2 - b1 * a2 - e1 * b2) / (a1 * b2 + b1 * c2 + c1 * a2 - a1 * c2 - b1 * a2 - c1 * b2);

            if (RGB)
            {
                Lw = 0;
                α = 0;
                β = 0;
                γ = 0;
            }
            else
            {
                if ((L1_1 >= 0) && (L1_2 >= 0))
                {
                    Lw = L1_3;
                    α = 0;
                    β = Math.Round(L1_1 / Lg, 4);
                    γ = Math.Round(L1_2 / Lb, 4);
                }
                else if ((L2_1 >= 0) && (L2_2 >= 0))
                {
                    Lw = L2_3;
                    α = Math.Round(L2_1 / Lr, 4);
                    β = 0;
                    γ = Math.Round(L2_2 / Lb, 4);
                }
                else if ((L3_1 >= 0) && (L3_2 >= 0))
                {
                    Lw = L3_3;
                    α = Math.Round(L3_2 / Lr, 4);
                    β = Math.Round(L3_1 / Lg, 4);
                    γ = 0;
                }
                else
                {
                    MessageBox.Show("计算错误！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        public void LUTGen()
        {
            double[] r_trg = new double[2048];
            double[] g_trg = new double[2048];
            double[] b_trg = new double[2048];
            double[] w_trg = new double[2048];
            double[] lum = new double[2048];
            double[] r_ext = new double[4096];
            double[] g_ext = new double[4096];
            double[] b_ext = new double[4096];
            double[] w_ext = new double[4096];
            int[] lut_r = new int[2048];
            int[] lut_g = new int[2048];
            int[] lut_b = new int[2048];
            int[] lut_w = new int[2048];

            LrLgLbLwCal();

            for (int i = 0; i < addresswidth; i++)
            {
                lum[i] = tarlum * Math.Pow(i, gamma) / Math.Pow((addresswidth - 1), gamma);
                r_trg[i] = lum[i] * Lr;
                g_trg[i] = lum[i] * Lg;
                b_trg[i] = lum[i] * Lb;
                if (RGB)
                {
                    w_trg[i] = 0;
                }
                else
                {
                    w_trg[i] = lum[i] * Lw;
                }
            }

            for (int i = 0; i < 256; i++)
            {
                for(int j = 0; j < (datawidth / 256); j++)
                {
                    if (i == 255)
                    {
                        r_ext[(datawidth / 256) * i + j] = R_LUM[i];
                        g_ext[(datawidth / 256) * i + j] = G_LUM[i];
                        b_ext[(datawidth / 256) * i + j] = B_LUM[i];
                        w_ext[(datawidth / 256) * i + j] = W_LUM[i];
                    }
                    else
                    {
                        r_ext[(datawidth / 256) * i + j] = R_LUM[i] + (R_LUM[i + 1] - R_LUM[i]) / (datawidth / 256) * j;
                        g_ext[(datawidth / 256) * i + j] = G_LUM[i] + (G_LUM[i + 1] - G_LUM[i]) / (datawidth / 256) * j;
                        b_ext[(datawidth / 256) * i + j] = B_LUM[i] + (B_LUM[i + 1] - B_LUM[i]) / (datawidth / 256) * j;
                        w_ext[(datawidth / 256) * i + j] = W_LUM[i] + (W_LUM[i + 1] - W_LUM[i]) / (datawidth / 256) * j;
                    }
                }
            }

            for (int i = 0; i < addresswidth; i++)
            {
                if (r_trg[i] <= r_ext[0])
                {
                    lut_r[i] = 0;
                }
                else if (r_trg[i] >= r_ext[datawidth - 1])
                {
                    lut_r[i] = datawidth - 1;
                }
                else
                {
                    for (int j = 0; j < (datawidth - 1); j++)
                    {
                        if ((r_trg[i] >= r_ext[j]) && (r_trg[i] < r_ext[j + 1]))
                        {
                            lut_r[i] = j;
                            break;
                        }
                    }
                }
            }

            for (int i = 0; i < addresswidth; i++)
            {
                if (g_trg[i] <= g_ext[0])
                {
                    lut_g[i] = 0;
                }
                else if (g_trg[i] >= g_ext[datawidth - 1])
                {
                    lut_g[i] = datawidth - 1;
                }
                else
                {
                    for (int j = 0; j < (datawidth - 1); j++)
                    {
                        if ((g_trg[i] >= g_ext[j]) && (g_trg[i] < g_ext[j + 1]))
                        {
                            lut_g[i] = j;
                            break;
                        }
                    }
                }
            }

            for (int i = 0; i < addresswidth; i++)
            {
                if (b_trg[i] <= b_ext[0])
                {
                    lut_b[i] = 0;
                }
                else if (b_trg[i] >= b_ext[datawidth - 1])
                {
                    lut_b[i] = datawidth - 1;
                }
                else
                {
                    for (int j = 0; j < (datawidth - 1); j++)
                    {
                        if ((b_trg[i] >= b_ext[j]) && (b_trg[i] < b_ext[j + 1]))
                        {
                            lut_b[i] = j;
                            break;
                        }
                    }
                }
            }

            for (int i = 0; i < addresswidth; i++)
            {
                if (w_trg[i] <= w_ext[0])
                {
                    lut_w[i] = 0;
                }
                else if (w_trg[i] >= w_ext[datawidth - 1])
                {
                    lut_w[i] = datawidth - 1;
                }
                else
                {
                    for (int j = 0; j < (datawidth - 1); j++)
                    {
                        if ((w_trg[i] >= w_ext[j]) && (w_trg[i] < w_ext[j + 1]))
                        {
                            lut_w[i] = j;
                            break;
                        }
                    }
                }
            }

            if (addresswidth == 256)
            {
                for (int i = 0; i < addresswidth; i++)
                {
                    flash[4 * i] = (byte)(lut_r[i] >> 8);
                    flash[4 * i + 1] = (byte)(lut_r[i] - flash[4 * i] * 256);
                    flash[4 * i + 2] = (byte)(lut_g[i] >> 8);
                    flash[4 * i + 3] = (byte)(lut_g[i] - flash[4 * i + 2] * 256);
                    flash[4 * i + 1024] = (byte)(lut_b[i] >> 8);
                    flash[4 * i + 1025] = (byte)(lut_b[i] - flash[4 * i + 1024] * 256);
                    if (RGB)
                    {
                        flash[4 * i + 1026] = 0;
                        flash[4 * i + 1027] = 0;
                    }
                    else
                    {
                        flash[4 * i + 1026] = (byte)(lut_w[i] >> 8);
                        flash[4 * i + 1027] = (byte)(lut_w[i] - flash[4 * i + 1026] * 256);
                    }
                }
            }
            else
            {
                for (int i = 0; i < addresswidth; i++)
                {
                    flash[2 * i] = (byte)(lut_r[i] % 256);
                    flash[2 * i + 1] = (byte)(lut_r[i] / 256);
                    flash[2 * i + addresswidth * 2] = (byte)(lut_g[i] % 256);
                    flash[2 * i + 1 + addresswidth * 2] = (byte)(lut_g[i] / 256);
                    flash[2 * i + addresswidth * 4] = (byte)(lut_b[i] % 256);
                    flash[2 * i + 1 + addresswidth * 4] = (byte)(lut_b[i] / 256);
                    if (RGB)
                    {
                        flash[2 * i + addresswidth * 6] = 0;
                        flash[2 * i + 1 + addresswidth * 6] = 0;
                    }
                    else
                    {
                        flash[2 * i + addresswidth * 6] = (byte)(lut_w[i] % 256);
                        flash[2 * i + 1 + addresswidth * 6] = (byte)(lut_w[i] / 256);
                    }
                }
            }
            flash[addresswidth * 8] = (byte)(α * 65536 % 256);
            flash[addresswidth * 8 + 1] = (byte)(α * 65536 / 256);
            flash[addresswidth * 8 + 2] = (byte)(β * 65536 % 256);
            flash[addresswidth * 8 + 3] = (byte)(β * 65536 / 256);
            flash[addresswidth * 8 + 4] = (byte)(γ * 65536 % 256);
            flash[addresswidth * 8 + 5] = (byte)(γ * 65536 / 256);
        }

    }
}
