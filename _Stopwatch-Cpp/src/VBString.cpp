#include "VBString.h"
#include <string>
#include <sstream>

using namespace std;

string VBString::StringS(int len, string c)
{
    std::stringstream ss;
    for (int i=0; i<len; ++i)
    {
        ss << c;
    }
    return ss.str();
}
string VBString::CStr(unsigned char c)
{
    std::stringstream s; s << c; return s.str();
}
string VBString::CStr(char c)
{
    std::stringstream s; s << c; return s.str();
}
string VBString::CStr(unsigned short i)
{
    std::stringstream s; s << i; return s.str();
}
string VBString::CStr(short i)
{
    std::stringstream s; s << i; return s.str();
}
string VBString::CStr(unsigned int i)
{
    std::stringstream s; s << i; return s.str();
}
string VBString::CStr(int i)
{
    std::stringstream s; s << i; return s.str();
}
string VBString::CStr(unsigned long i)
{
    std::stringstream s; s << i; return s.str();
}
string VBString::CStr(long i)
{
    std::stringstream s; s << i; return s.str();
}
string VBString::CStr(unsigned long long i)
{
    std::stringstream s; s << i; return s.str();
}
string VBString::CStr(long long i)
{
    std::stringstream s; s << i; return s.str();
}
string VBString::CStr(float f) //aka Single
{
    std::stringstream s; s << f; return s.str();
}
string VBString::CStr(double f)
{
    std::stringstream s; s << f; return s.str();
}
string VBString::CStr(long double f)
{
    std::stringstream s; s << f; return s.str();
}
string VBString::CStr(bool b)
{
    return (b == 1) ? ("true") : ("false");
}
string VBString::IntToString(int n, int digits)
{
    string s = CStr(n);
    int l = s.length();
    if (l<digits)
    {
        s = StringS(digits - l, "0") + s;
    }
    return s;
}

