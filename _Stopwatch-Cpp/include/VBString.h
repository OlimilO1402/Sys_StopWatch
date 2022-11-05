#ifndef VBSTRING_H
#define VBSTRING_H
#include <string>
using namespace std;
namespace VBString
{
    string StringS(int len, string c);
    string CStr(unsigned char c);
    string CStr(char c);
    string CStr(unsigned short i);
    string CStr(short i);
    string CStr(unsigned int i);
    string CStr(int i);
    string CStr(unsigned long i);
    string CStr(long i);
    string CStr(unsigned long long i);
    string CStr(long long i);
    string CStr(float f); //aka Single
    string CStr(double f);
    string CStr(long double f);
    string CStr(bool b);
    string IntToString(int n, int digits);
}
#endif // VBSTRING_H
