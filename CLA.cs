using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OC_Porject_20191112
{
    class CLA
    {
        public byte[] readflash = new byte[16390];
        public byte[] writeflash = new byte[32774];
        public DataTable dt_CLA;
        public double Imax;

        double[] cla_r = new double[2048];
        double[] cla_g = new double[2048];
        double[] cla_b = new double[2048];
        double[] cla_w = new double[2048];
        int[] point = { 0, 240, 481, 721, 961, 1202, 1442, 1683, 1915 };

        public void LUTRead(byte[] readflash)
        {
            for (int i = 0; i < 16390; i++)
            {
                writeflash[i + 2048 * 8] = readflash[i];
            }
        }

        public void CLAcal()
        {
            for(int i = 0; i < 8; i++)
            {
                int p2 = point[i + 1];
                int p1 = point[i];
                double R_I2 = Convert.ToDouble(dt_CLA.Rows[i + 2][1]);
                double R_I1 = Convert.ToDouble(dt_CLA.Rows[i + 1][1]);
                double G_I2 = Convert.ToDouble(dt_CLA.Rows[i + 2][2]);
                double G_I1 = Convert.ToDouble(dt_CLA.Rows[i + 1][2]);
                double B_I2 = Convert.ToDouble(dt_CLA.Rows[i + 2][3]);
                double B_I1 = Convert.ToDouble(dt_CLA.Rows[i + 1][3]);
                double W_I2 = Convert.ToDouble(dt_CLA.Rows[i + 2][4]);
                double W_I1 = Convert.ToDouble(dt_CLA.Rows[i + 1][4]);
                if(i == 7)
                {
                    for (int j = point[i]; j < 2048; j++)
                    {
                        cla_r[j] = ((R_I2 - R_I1) * j + (R_I1 * p2 - R_I2 * p1)) / (p2 - p1);
                        cla_g[j] = ((G_I2 - G_I1) * j + (G_I1 * p2 - G_I2 * p1)) / (p2 - p1);
                        cla_b[j] = ((B_I2 - B_I1) * j + (B_I1 * p2 - B_I2 * p1)) / (p2 - p1);
                        cla_w[j] = ((W_I2 - W_I1) * j + (W_I1 * p2 - W_I2 * p1)) / (p2 - p1);
                    }
                }
                else
                {
                    for (int j = point[i]; j < point[i + 1]; j++)
                    {
                        cla_r[j] = ((R_I2 - R_I1) * j + (R_I1 * p2 - R_I2 * p1)) / (p2 - p1);
                        cla_g[j] = ((G_I2 - G_I1) * j + (G_I1 * p2 - G_I2 * p1)) / (p2 - p1);
                        cla_b[j] = ((B_I2 - B_I1) * j + (B_I1 * p2 - B_I2 * p1)) / (p2 - p1);
                        cla_w[j] = ((W_I2 - W_I1) * j + (W_I1 * p2 - W_I2 * p1)) / (p2 - p1);
                    }
                }
            }
        }

        public void LUTWrite()
        {
            CLAcal();
            for (int i = 0; i < 2048; i++)
            {
                writeflash[2 * i] = (byte)(cla_r[i] /Imax * 65536 % 256);
                writeflash[2 * i + 1] = (byte)(cla_r[i] / Imax * 65536 / 256);
                writeflash[2 * i + 2048 * 2] = (byte)(cla_g[i] / Imax * 65536 % 256);
                writeflash[2 * i + 1 + 2048 * 2] = (byte)(cla_g[i] / Imax * 65536 / 256);
                writeflash[2 * i + 2048 * 4] = (byte)(cla_b[i] / Imax * 65536 % 256);
                writeflash[2 * i + 1 + 2048 * 4] = (byte)(cla_b[i] / Imax * 65536 / 256);
                writeflash[2 * i + 2048 * 6] = (byte)(cla_w[i] / Imax * 65536 % 256);
                writeflash[2 * i + 1 + 2048 * 6] = (byte)(cla_w[i] / Imax * 65536 / 256);
            }
        }

    }
}
