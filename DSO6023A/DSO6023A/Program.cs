/*
* Agilent VISA COM Example in C#
* -------------------------------------------------------------------
* This program illustrates most of the commonly used programming
* features of your Agilent oscilloscopes.
* -------------------------------------------------------------------
*/
using System;
using System.IO;
using System.Text;
using Ivi.Visa.Interop;
using System.Runtime.InteropServices;
namespace InfiniiVision
{
    class VisaComInstrumentApp
    {
    private static VisaComInstrument myScope;
    public static void Main(string[] args)
    {
        try
        {
            myScope = new
            VisaComInstrument("USB0::0x0957::0x1732::MY55270124::INSTR");
            //Initialize();
            /* The extras function contains miscellaneous commands that
            * do not need to be executed for the proper operation of
            * this example. The commands in the extras function are
            * shown for reference purposes only.
            */
            //Extra(); // Uncomment to execute the extra function.
            Capture();
            Console.WriteLine("Enter file name:");
            string fName = Console.ReadLine();
            Analyze(fName); //input fName
        }
        catch (System.ApplicationException err)
        {
            Console.WriteLine("*** VISA Error Message : " + err.Message);
        }
        catch (System.SystemException err)
        {
            Console.WriteLine("*** System Error Message : " + err.Message);
        }
        catch (System.Exception err)
        {
            System.Diagnostics.Debug.Fail("Unexpected Error");
            Console.WriteLine("*** Unexpected Error : " + err.Message);
        }
        finally
        {
             myScope.Close();
        }
        Console.WriteLine("Press enter to close..."); //time delay
        Console.ReadLine();
    }
/*
* Initialize()
* --------------------------------------------------------------
* This function initializes both the interface and the
* oscilloscope to a known state.
*/
//private static void Initialize()
//{
//string strResults;
///* RESET - This command puts the oscilloscope into a known
//* state. This statement is very important for programs to
//* work as expected. Most of the following initialization
//* commands are initialized by *RST. It is not necessary to
//* reinitialize them unless the default setting is not suitable
//* for your application.
//*/
//myScope.DoCommand("*RST"); // Reset the to the defaults.
//myScope.DoCommand("*CLS"); // Clear the status data structures.
///* IDN - Ask for the device's *IDN string.
//*/
//strResults = myScope.DoQueryString("*IDN?");
//// Display results.
//Console.Write("Result is: {0}", strResults);
///* AUTOSCALE - This command evaluates all the input signals
//* and sets the correct conditions to display all of the
//* active signals.
//*/
//myScope.DoCommand(":AUToscale");
///* CHANNEL_PROBE - Sets the probe attenuation factor for the
//* selected channel. The probe attenuation factor may be from
//* 0.1 to 1000.
//*/
//myScope.DoCommand(":CHANnel1:PROBe 1");
///* CHANNEL_RANGE - Sets the full scale vertical range in volts.
//* The range value is eight times the volts per division.
//*/
//myScope.DoCommand(":CHANnel1:RANGe 8");
///* TIME_RANGE - Sets the full scale horizontal time in seconds.
//* The range value is ten times the time per division.
//*/
//myScope.DoCommand(":TIMebase:RANGe 2e-3");
///* TIME_REFERENCE - Possible values are LEFT and CENTER:
//* - LEFT sets the display reference one time division from
//* the left.
//* - CENTER sets the display reference to the center of the
//* screen.
//*/
//myScope.DoCommand(":TIMebase:REFerence CENTer");
///* TRIGGER_SOURCE - Selects the channel that actually produces
//* the TV trigger. Any channel can be selected.
//*/
//myScope.DoCommand(":TRIGger:TV:SOURCe CHANnel1");
///* TRIGGER_MODE - Set the trigger mode to, EDGE, GLITch,
//* PATTern, CAN, DURation, IIC, LIN, SEQuence, SPI, TV,
//* UART, or USB.
//*/
//myScope.DoCommand(":TRIGger:MODE EDGE");
///* TRIGGER_EDGE_SLOPE - Set the slope of the edge for the
//* trigger to either POSITIVE or NEGATIVE.
//*/
//myScope.DoCommand(":TRIGger:EDGE:SLOPe POSitive");
//}
///*
//* Extra()
//* --------------------------------------------------------------
//* The commands in this function are not executed and are shown
//* for reference purposes only. To execute these commands, call
//* this function from main.
//*/
//private static void Extra()
//{
///* RUN_STOP (not executed in this example):
//* - RUN starts the acquisition of data for the active
//* waveform display.
//* - STOP stops the data acquisition and turns off AUTOSTORE.
//*/
//myScope.DoCommand(":RUN");
//myScope.DoCommand(":STOP");
///* VIEW_BLANK (not executed in this example):
//* - VIEW turns on (starts displaying) an active channel or
//* pixel memory.
//* - BLANK turns off (stops displaying) a specified channel or
//* pixel memory.
//*/
//myScope.DoCommand(":BLANk CHANnel1");
//myScope.DoCommand(":VIEW CHANnel1");
///* TIME_MODE (not executed in this example) - Set the time base
//* mode to MAIN, DELAYED, XY or ROLL.
//*/
//myScope.DoCommand(":TIMebase:MODE MAIN");
//}
///*
//* Capture()
//* --------------------------------------------------------------
//* This function prepares the scope for data acquisition and then
//* uses the DIGITIZE MACRO to capture some data.
//
private static void Capture()
{
/* AQUIRE_TYPE - Sets the acquisition mode. There are three
* acquisition types NORMAL, PEAK, or AVERAGE.
*/
myScope.DoCommand(":ACQuire:TYPE NORMal");
/* AQUIRE_COMPLETE - Specifies the minimum completion criteria
* for an acquisition. The parameter determines the percentage
* of time buckets needed to be "full" before an acquisition is
* considered to be complete.
*/
myScope.DoCommand(":ACQuire:COMPlete 100");
/* DIGITIZE - Used to acquire the waveform data for transfer
* over the interface. Sending this command causes an
* acquisition to take place with the resulting data being
* placed in the buffer.
*/
/* NOTE! The use of the DIGITIZE command is highly recommended
* as it will ensure that sufficient data is available for
* measurement. Keep in mind when the oscilloscope is running,
* communication with the computer interrupts data acquisition.
* Setting up the oscilloscope over the bus causes the data
* buffers to be cleared and internal hardware to be
* reconfigured.
* If a measurement is immediately requested there may not have
* been enough time for the data acquisition process to collect
* data and the results may not be accurate. An error value of
* 9.9E+37 may be returned over the bus in this situation.
*/
myScope.DoCommand(":DIGitize CHANnel1");
}
/*
* Analyze()
* --------------------------------------------------------------
* In this example we will do the following:
* - Save the system setup to a file for restoration at a later
* time.
* - Save the oscilloscope display to a file which can be
* printed.
* - Make single channel measurements.
*/
private static void Analyze(string fName)
{
byte[] ResultsArray; // Results array.
int nBytes; // Number of bytes returned from instrument.
/* SAVE_SYSTEM_SETUP - The :SYSTem:SETup? query returns a
* program message that contains the current state of the
* instrument. Its format is a definite-length binary block,
* for example,
* #800002204<setup string><NL>
* where the setup string is 2204 bytes in length.
*/
//-----------------------------------------------------------------------------------
//Console.WriteLine("Saving oscilloscope setup to " +
//"c:\\scope\\config\\setup.dat");
//if (File.Exists("c:\\scope\\config\\setup.dat"))
//File.Delete("c:\\scope\\config\\setup.dat");
//// Query and read setup string.
//ResultsArray = myScope.DoQueryIEEEBlock(":SYSTem:SETup?");
//nBytes = ResultsArray.Length;
//Console.WriteLine("Read oscilloscope setup ({0} bytes).",
//nBytes);
//// Write setup string to file.
//File.WriteAllBytes("c:\\scope\\config\\setup.dat",
//ResultsArray);
//Console.WriteLine("Wrote setup string ({0} bytes) to file.",
//nBytes);
///* RESTORE_SYSTEM_SETUP - Uploads a previously saved setup
//* string to the oscilloscope.
//*/
//byte[] DataArray;
//// Read setup string from file.
//DataArray = File.ReadAllBytes("c:\\scope\\config\\setup.dat");
//Console.WriteLine("Read setup string ({0} bytes) from file.",
//DataArray.Length);
//// Restore setup string.
//myScope.DoCommandIEEEBlock(":SYSTem:SETup", DataArray);
//Console.WriteLine("Restored setup string.");
///* IMAGE_TRANSFER - In this example, we query for the screen
//* data with the ":DISPLAY:DATA?" query. The .png format
//* data is saved to a file in the local file system.
//*/
//Console.WriteLine("Transferring screen image to " +
//"c:\\scope\\data\\screen.png");
//if (File.Exists("c:\\scope\\data\\screen.png"))
//File.Delete("c:\\scope\\data\\screen.png");
//// Increase I/O timeout to fifteen seconds.
//myScope.SetTimeoutSeconds(15);
//// Get the screen data in PNG format.
//ResultsArray = myScope.DoQueryIEEEBlock(
//":DISPlay:DATA? PNG, SCReen, COLor");
//nBytes = ResultsArray.Length;
//Console.WriteLine("Read screen image ({0} bytes).", nBytes);
//// Store the screen data in a file.
//File.WriteAllBytes("c:\\scope\\data\\screen.png",
//ResultsArray);
//Console.WriteLine("Wrote screen image ({0} bytes) to file.",
//nBytes);
//// Return I/O timeout to five seconds.
//myScope.SetTimeoutSeconds(5);
//-----------------------------------------------------------------------------------
/* MEASURE - The commands in the MEASURE subsystem are used to
* make measurements on displayed waveforms.
*/
// Set source to measure.
myScope.DoCommand(":MEASure:SOURce CHANnel1");
// Query for frequency.
double fResults;
fResults = myScope.DoQueryValue(":MEASure:FREQuency?");
Console.WriteLine("The frequency is: {0:F4} kHz",
fResults / 1000);
// Query for peak to peak voltage.
fResults = myScope.DoQueryValue(":MEASure:VPP?");
Console.WriteLine("The peak to peak voltage is: {0:F2} V",
fResults);
/* WAVEFORM_DATA - Get waveform data from oscilloscope. To
844 Agilent InfiniiVision 6000 Series Oscilloscopes Programmer's Guide
12 Programming Examples
* obtain waveform data, you must specify the WAVEFORM
* parameters for the waveform data prior to sending the
* ":WAVEFORM:DATA?" query.
*
* Once these parameters have been sent, the
* ":WAVEFORM:PREAMBLE?" query provides information concerning
* the vertical and horizontal scaling of the waveform data.
*
* With the preamble information you can then use the
* ":WAVEFORM:DATA?" query and read the data block in the
* correct format.
*/
/* WAVE_FORMAT - Sets the data transmission mode for waveform
* data output. This command controls how the data is
* formatted when sent from the oscilloscope and can be set
* to WORD or BYTE format.
*/
// Set waveform format to BYTE.
myScope.DoCommand(":WAVeform:FORMat BYTE");


 /// Print maximum point over measurement
 //double [] maxPoint = myScope.DoQueryValues("ACQUIRE:POINTS?");
 //foreach (var item in maxPoint)
 //{
 //    Console.WriteLine("Maximal points: {0:e}", item); ///4e6 points
 //}

/* WAVE_POINTS - Sets the number of points to be transferred.
* The number of time points available is returned by the
* "ACQUIRE:POINTS?" query. This can be set to any binary
* fraction of the total time points available.
*/
myScope.DoCommand(":WAVeform:POINts 1000");
/* GET_PREAMBLE - The preamble contains all of the current
* WAVEFORM settings returned in the form <preamble block><NL>
* where the <preamble block> is:
* FORMAT : int16 - 0 = BYTE, 1 = WORD, 4 = ASCII.
* TYPE : int16 - 0 = NORMAL, 1 = PEAK DETECT,
* 2 = AVERAGE.
* POINTS : int32 - number of data points transferred.
* COUNT : int32 - 1 and is always 1.
* XINCREMENT : float64 - time difference between data
* points.
* XORIGIN : float64 - always the first data point in
* memory.
* XREFERENCE : int32 - specifies the data point associated
* with the x-origin.
* YINCREMENT : float32 - voltage difference between data
* points.
* YORIGIN : float32 - value of the voltage at center
* screen.
* YREFERENCE : int32 - data point where y-origin occurs.
*/
Console.WriteLine("Reading preamble.");
double[] fResultsArray;
fResultsArray = myScope.DoQueryValues(":WAVeform:PREamble?");
double fFormat = fResultsArray[0];
Console.WriteLine("Preamble FORMat: {0:e}", fFormat);
double fType = fResultsArray[1];
Console.WriteLine("Preamble TYPE: {0:e}", fType);
double fPoints = fResultsArray[2];
Console.WriteLine("Preamble POINts: {0:e}", fPoints);
double fCount = fResultsArray[3];
Console.WriteLine("Preamble COUNt: {0:e}", fCount);
double fXincrement = fResultsArray[4];
Console.WriteLine("Preamble XINCrement: {0:e}", fXincrement);
double fXorigin = fResultsArray[5];
Console.WriteLine("Preamble XORigin: {0:e}", fXorigin);
double fXreference = fResultsArray[6];
Console.WriteLine("Preamble XREFerence: {0:e}", fXreference);
double fYincrement = fResultsArray[7];
Console.WriteLine("Preamble YINCrement: {0:e}", fYincrement);
double fYorigin = fResultsArray[8];
Console.WriteLine("Preamble YORigin: {0:e}", fYorigin);
double fYreference = fResultsArray[9];
Console.WriteLine("Preamble YREFerence: {0:e}", fYreference);
/* QUERY_WAVE_DATA - Outputs waveform records to the controller
* over the interface that is stored in a buffer previously
* specified with the ":WAVeform:SOURce" command.
*/
/* READ_WAVE_DATA - The wave data consists of two parts: the
* header, and the actual waveform data followed by a
* New Line (NL) character. The query data has the following
* format:
*
* <header><waveform data block><NL>
*
* Where:
*
* <header> = #800002048 (this is an example header)
*
* The "#8" may be stripped off of the header and the remaining
* numbers are the size, in bytes, of the waveform data block.
* The size can vary depending on the number of points acquired
* for the waveform which can be set using the
* ":WAVEFORM:POINTS" command. You may then read that number
* of bytes from the oscilloscope; then, read the following NL
* character to terminate the query.
*/
// Read waveform data.
ResultsArray = myScope.DoQueryIEEEBlock(":WAVeform:DATA?");
nBytes = ResultsArray.Length;
Console.WriteLine("Read waveform data ({0} bytes).", nBytes);
// Make some calculations from the preamble data.
double fVdiv = 32 * fYincrement;
double fOffset = fYorigin;
double fSdiv = fPoints * fXincrement / 10;
double fDelay = (fPoints / 2) * fXincrement + fXorigin;
// Print them out...
Console.WriteLine("Scope Settings for Channel 1:");
Console.WriteLine("Volts per Division = {0:f}", fVdiv);
Console.WriteLine("Offset = {0:f}", fOffset);
Console.WriteLine("Seconds per Division = {0:e}", fSdiv);
Console.WriteLine("Delay = {0:e}", fDelay);
// Print the waveform voltage at selected points:
for (int i = 0; i < nBytes; i = i + (nBytes / 20))
{
Console.WriteLine("Data point {0:d} = {1:f6} Volts at "
+ "{2:f10} Seconds", i,
((float)ResultsArray[i] - fYreference) * fYincrement +
fYorigin,
((float)i - fXreference) * fXincrement + fXorigin);
}
/* SAVE_WAVE_DATA - saves the waveform data to a CSV format
* file named "waveform.csv".
*/
if (File.Exists("c:\\scope\\data\\" + fName + ".csv"))  
    File.Delete("c:\\scope\\data\\" + fName + ".csv");
StreamWriter writer =
File.CreateText("c:\\scope\\data\\" + fName + ".csv");
for (int i = 0; i < nBytes; i++)
{
writer.WriteLine("{0:E}, {1:f6}",
((float)i - fXreference) * fXincrement + fXorigin,
((float)ResultsArray[i] - fYreference) * fYincrement +
fYorigin);
}
writer.Close();
Console.WriteLine("Waveform data ({0} points) written to " +
"c:\\scope\\data\\" + fName + ".csv.", nBytes);    //"c:\\scope\\data\\waveform.csv."
}
}
class VisaComInstrument
{
private ResourceManagerClass m_ResourceManager;
private FormattedIO488Class m_IoObject;
private string m_strVisaAddress;
// Constructor.
public VisaComInstrument(string strVisaAddress)
{
// Save VISA address in member variable.
m_strVisaAddress = strVisaAddress;
// Open the default VISA COM IO object.
OpenIo();
// Clear the interface.
m_IoObject.IO.Clear();
}
public void DoCommand(string strCommand)
{
// Send the command.
m_IoObject.WriteString(strCommand, true);
// Check for instrument errors.
CheckForInstrumentErrors(strCommand);
}
public string DoQueryString(string strQuery)
{
// Send the query.
m_IoObject.WriteString(strQuery, true);
// Get the result string.
string strResults;
strResults = m_IoObject.ReadString();
// Check for instrument errors.
CheckForInstrumentErrors(strQuery);
// Return results string.
return strResults;
}
public double DoQueryValue(string strQuery)
{
// Send the query.
m_IoObject.WriteString(strQuery, true);
// Get the result number.
double fResult;
fResult = (double)m_IoObject.ReadNumber(
IEEEASCIIType.ASCIIType_R8, true);
// Check for instrument errors.
CheckForInstrumentErrors(strQuery);
// Return result number.
return fResult;
}
public double[] DoQueryValues(string strQuery)
{
// Send the query.
m_IoObject.WriteString(strQuery, true);
// Get the result numbers.
double[] fResultsArray;
fResultsArray = (double[])m_IoObject.ReadList(
IEEEASCIIType.ASCIIType_R8, ",;");

// Check for instrument errors.
CheckForInstrumentErrors(strQuery);
// Return result numbers.
return fResultsArray;
}
public byte[] DoQueryIEEEBlock(string strQuery)
{
// Send the query.
m_IoObject.WriteString(strQuery, true);
// Get the results array.
byte[] ResultsArray;
ResultsArray = (byte[])m_IoObject.ReadIEEEBlock(
IEEEBinaryType.BinaryType_UI1, false, true);
// Check for instrument errors.
CheckForInstrumentErrors(strQuery);
// Return results array.
return ResultsArray;
}
public void DoCommandIEEEBlock(string strCommand,
byte[] DataArray)
{
// Send the command.
m_IoObject.WriteIEEEBlock(strCommand, DataArray, true);
// Check for instrument errors.
CheckForInstrumentErrors(strCommand);
}
private void CheckForInstrumentErrors(string strCommand)
{
string strInstrumentError;
bool bFirstError = true;
// Repeat until all errors are displayed.
do
{
// Send the ":SYSTem:ERRor?" query, and get the result string.
m_IoObject.WriteString(":SYSTem:ERRor?", true);
strInstrumentError = m_IoObject.ReadString();
// If there is an error, print it.
if (strInstrumentError.ToString() != "+0,\"No error\"\n")
{
if (bFirstError)
{
// Print the command that caused the error.
Console.WriteLine("ERROR(s) for command '{0}': ",
strCommand);
bFirstError = false;
}
Console.Write(strInstrumentError);

}
} while (strInstrumentError.ToString() != "+0,\"No error\"\n");
}
private void OpenIo()
{
m_ResourceManager = new ResourceManagerClass();
m_IoObject = new FormattedIO488Class();
// Open the default VISA COM IO object.
try
{
m_IoObject.IO =
(IMessage)m_ResourceManager.Open(m_strVisaAddress,
AccessMode.NO_LOCK, 0, "");
}
catch (Exception e)
{
Console.WriteLine("An error occurred: {0}", e.Message);
}
}
public void SetTimeoutSeconds(int nSeconds)
{
m_IoObject.IO.Timeout = nSeconds * 1000;
}
public void Close()
{
try
{
m_IoObject.IO.Close();
}
catch {}
try
{
Marshal.ReleaseComObject(m_IoObject);
}
catch {}
try
{
Marshal.ReleaseComObject(m_ResourceManager);
}
catch {}
}
}
}

