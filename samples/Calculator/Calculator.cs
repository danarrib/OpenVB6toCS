public class Calculator
{
    // A simple calculator class for testing the translator
    private double mLastResult;

    public enum MathOperation
    {
        opAdd = 1,
        opSubtract = 2,
        opMultiply = 3,
        opDivide = 4,
    }

    public double Calculate(double a, double b, MathOperation op)
    {
        double _result = default;
        double result;
        switch (op)
        {
            case opAdd:
            {
                result = (a + b);
                break;
            }
            case opSubtract:
            {
                result = (a - b);
                break;
            }
            case opMultiply:
            {
                result = (a * b);
                break;
            }
            case opDivide:
            {
                if ((b == 0))
                {
                    throw new Exception("Division by zero"); // Err.Raise
                }
                result = (a / b);
                break;
            }
        }
        mLastResult = result;
        _result = result;
        return _result;
    }

    public double LastResult
    {
        get
        {
            double _result = default;
            _result = mLastResult;
            return _result;
        }
    }

    public void Reset()
    {
        mLastResult = 0;
    }
}
