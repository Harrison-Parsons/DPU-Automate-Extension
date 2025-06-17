using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UtilitiesAutomateExtension.Pages;
public class Person
{
    private string _name;
    private string _email;
    private string _phoneNumber;

    public Person(string name, string email, string phoneNumber)
    {
        _name = name;
        _email = email;
        _phoneNumber = phoneNumber;
    }

    public string getName()
    {
        return _name;
    }

    public string getEmail()
    {
        return _email;
    }

    public string getPhoneNumber()
    {
        return _phoneNumber;
    }

    public override string ToString()
    {
        return $"{_name} - {_phoneNumber}";
    }
}