#include "Stopwatch.h"
#include <windows.h>
#include <iostream>
#include <string>
#include <sstream>
#include <VBString.h>
#include <ctime>
#define TicksPerMillisecond 10000
//const unsigned long TicksPerMillisecond = 10000;
#define TicksPerSecond      10000000
//const unsigned long TicksPerSecond      = 10000000;

using namespace std;
using namespace VBString;

//public:
//ctor
Stopwatch::Stopwatch()
{
    LARGE_INTEGER frequency;
    BOOL b = QueryPerformanceFrequency(&frequency);
    m_Frequency = (unsigned long long)frequency.QuadPart;
    if (b == 0) {
        m_IsHighResolution = false;
        m_IsRunning = false;
        m_Frequency = TicksPerSecond;
        m_TickFrequency = 1;
    } else {
        m_IsHighResolution = true;
        m_TickFrequency = (TicksPerSecond / m_Frequency);
    }
}
unsigned long long Stopwatch::Timestamp()
{
    if (m_IsHighResolution){
        LARGE_INTEGER timestamp;
        BOOL b = QueryPerformanceCounter(&timestamp);
        return ((b == 1)?(timestamp.QuadPart):(0));
        //if (b == 1) {
        //    return timestamp.QuadPart
        //} else {return 0}
    } else {
        return clock(); //DateTime.Now
    }
}
void Stopwatch::Start()
{
    if (!m_IsRunning) {
        m_StartTimeStamp = Timestamp();
    }
}
Stopwatch Stopwatch::StartNew()
{
    Stopwatch sw;
    sw.Start();
    return sw;
}
void Stopwatch::Stop()
{
    if (m_IsRunning)
    {
        m_Elapsed += Timestamp() - m_StartTimeStamp;
        m_IsRunning = false;
    }
}
void Stopwatch::Reset()
{
    m_Elapsed = 0;
    m_IsRunning = false;
    m_StartTimeStamp = 0;
}
unsigned long long Stopwatch::Frequency()
{
    return m_Frequency;
}
bool Stopwatch::IsHighResolution()
{
    return m_IsHighResolution;
}
string Stopwatch::ElapsedToString()
{
    return TimeSpanToString(ElapsedDateTimeTicks());
}
unsigned long long Stopwatch::ElapsedMilliseconds()
{
    return ElapsedDateTimeTicks() / TicksPerMillisecond;
}
unsigned long long Stopwatch::ElapsedTicks()
{
    return RawElapsedTicks();
}
//private:
string Stopwatch::TimeSpanToString(unsigned long long ticks)
{
    stringstream b;
    //b = "Dies ist ein String"
    unsigned long long h, m, s, n;
    int days = (ticks / 86400000);
    unsigned long long time = (ticks % 86400000);

    if (ticks < 0) {
        b << "-";
        days = -days;
        time = -time;
    }

    if (days != 0) {
        b << days << ".";
    }

    //Stunden
    h = ((time / 3600000) % 24);
    b << IntToString(h, 2) << ":";

    //Minuten
    m = (int)((time / 60000) % 60);
    b << IntToString(m, 2) << ":";

    //Sekunden
    s = (int)((time / 1000) % 60);
    b << IntToString(s, 2);

    n = (ticks - (h * 3600000)
               - (m * 60000)
               - (s * 1000)) * 10000; //TicksPerMillisecond

    if (n != 0) {
        b << "." << IntToString(n, 7);
    }
    return b.str();
}
unsigned long long Stopwatch::ElapsedDateTimeTicks()
{
    unsigned long long rawElapsedTicks = RawElapsedTicks();
    if (m_IsHighResolution) {
        return rawElapsedTicks * m_TickFrequency;
    } else {
        return rawElapsedTicks;
    }
}
unsigned long long Stopwatch::RawElapsedTicks()
{
    unsigned long long elapsed = m_Elapsed;
    if (m_IsRunning)
    {
        elapsed += Timestamp() - m_StartTimeStamp;
    }
    return elapsed;
}

//*/
