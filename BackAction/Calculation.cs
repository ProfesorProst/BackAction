using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BackAction
{
    public static class Calculation
    {
        const double redPc = 1.054571817E-34;
        public static double spectrum(double obMass, double obFrequn, double cavResFrequn, double photonNum, 
            double phaseNoise, double prTransFreq, double lightSpeed, double measurTime, double measurNum, double photonNumFluct = 1)
        {
            double numerator = Math.Pow(redPc, 3) / (4 * Math.PI * obMass * obFrequn);
            numerator = (Math.Sqrt(numerator) * lightSpeed * cavResFrequn) / measurTime * photonNumFluct;
            numerator = Math.Pow(numerator, 2);

            double secondPart = (2 * photonNum) / (Math.PI * Math.Pow(prTransFreq, 2) * phaseNoise * measurNum);

            return numerator * secondPart;
        }
    }
}
