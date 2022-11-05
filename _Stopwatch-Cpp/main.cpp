#include <iostream>
#include <string>
#include <sstream>
#include <VBString.h>
#include <Stopwatch.h>
#include <windows.h>

using namespace std;

using namespace VBString;

int main()
{
/*
    LARGE_INTEGER f;
    BOOL boo = QueryPerformanceFrequency(&f);
    unsigned long frequ = f.QuadPart;
    if (boo == 1)
    {
        //frequ = (unsigned long)f.QuadPart;
    }
    cout << frequ << endl;
*/
    Stopwatch sw;
    sw.Start();
    unsigned char  uc = 65; // = "A";
    char           ch = -127; // = "ü";
    unsigned short us = 65535;
    short          sh = -32768;
    unsigned int   ui = 4294967294u;
    int            ih = -2147483647;
    unsigned long  ul = 123456798;
    long           lg = 123456789;
    unsigned long long ull = 123456789123456u;
    long long      ll = 123456789;
    float          fl = 123.456;
    double         db = 789.123;
    bool           bl = true;
    string a = "Hallo Welt! ";
    string b = "Dies ist ein String: ";
    sw.Stop();
    cout << sw.ElapsedToString() << endl;
    sw.Reset();
    sw = sw.StartNew();
    //Variante 1
    string d = a + b + CStr(uc) + " " + CStr(ch) + " " + CStr(us) + " " + CStr(sh) + " "
                     + CStr(ui) + " " + CStr(ih) + " " + CStr(ul) + " " + CStr(lg) + " "
                     + CStr(ull) + " " + CStr(ll) + " " + CStr(fl) + " " + CStr(db) + " " + CStr(bl);
    cout << d << endl;
    sw.Stop();
    cout << sw.ElapsedToString() << endl;

    sw.Reset();
    sw = sw.StartNew();
    //Variante 2
    stringstream ss;
    ss << a << b << uc  << " " << ch << " " << us << " " << sh << " "
                 << ui  << " " << ih << " " << ul << " " << lg << " "
                 << ull << " " << ll << " " << fl << " " << db << " " << ((bl==true)?("true"):("false"));
    cout << ss.str() << endl;
    sw.Stop();
    cout << sw.ElapsedToString() << endl;

    return 0;
}


