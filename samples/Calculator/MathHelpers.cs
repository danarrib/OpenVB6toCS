public static class MathHelpers
{
    public static double Clamp(double value, double minVal, double maxVal)
    {
        double _result = default;
        if ((value < minVal))
        {
            _result = minVal;
        }
        else if ((value > maxVal))
        {
            _result = maxVal;
        }
        else
        {
            _result = value;
        }
        return _result;
    }

    public static double RoundTo(double value, int decimals)
    {
        double _result = default;
        double factor;
        factor = Math.Pow(10, decimals);
        _result = ((int)Math.Floor(((value * factor) + 0.5)) / factor);
        return _result;
    }
}
