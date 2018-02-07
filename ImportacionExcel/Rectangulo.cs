using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportacionExcel
{
    class Rectangulo
    {
        private double b { get; set; }
        private double h { get; set; }

        public Rectangulo(double b, double h)
        {
            this.b = b;
            this.h = h;
        }

        public double getArea()
        {
            return b * h;
        }

        public double getPerimetro()
        {
            return b * 2 + h * 2;
        }
    }
}
