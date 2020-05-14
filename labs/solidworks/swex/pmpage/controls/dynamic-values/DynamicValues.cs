using System;
using System.ComponentModel;

public class DynamicValuesDataModel : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;

    private double m_Val1;
    private double m_Val2;

    public double Val1
    {
        get
        {
            return m_Val1;
        }
        set
        {
            m_Val1 = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Val1)));

            Val2 = value * 2;
        }
    }

    public double Val2
    {
        get
        {
            return m_Val2;
        }
        set
        {
            m_Val2 = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Val2)));
        }
    }

    public Action Reset => OnResetClick;

    private void OnResetClick()
    {
        Val1 = 10;
    }
}
