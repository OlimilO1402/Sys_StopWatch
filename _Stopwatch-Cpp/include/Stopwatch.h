#include <iostream>
#include <string>
#include <sstream>

#ifndef STOPWATCH_H
#define STOPWATCH_H

using namespace std;

class Stopwatch
{
    public:
        //ctor
        Stopwatch();
        //functions
        unsigned long long Timestamp();
        void Start();
        Stopwatch StartNew();
        void Stop();
        void Reset();
        unsigned long long Frequency();
        bool IsHighResolution();
        string ElapsedToString();
        unsigned long long ElapsedMilliseconds();
        unsigned long long ElapsedTicks();

    private:
        //constants
        //static const unsigned long TicksPerMillisecond = 10000u;
        //static const unsigned long TicksPerSecond      = 10000000u;

        //variables
        unsigned long long m_Frequency;
        bool m_IsHighResolution;
        bool m_IsRunning;
        unsigned long long m_StartTimeStamp;
        unsigned long long m_Elapsed;
        unsigned long long m_TickFrequency;

        //functions
        string TimeSpanToString(unsigned long long ticks);
        unsigned long long ElapsedDateTimeTicks();
        unsigned long long RawElapsedTicks();
};

#endif // STOPWATCH_H
