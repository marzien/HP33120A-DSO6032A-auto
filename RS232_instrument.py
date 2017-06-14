import visa

class agilent33120A:
   
    def __init__(self, device):
        rm = visa.ResourceManager()
        rm.list_resources()
        self.meas = rm.open_resource(device)

    def control(self, control):
        if control.find('local') != -1:
            self.meas.write("SYSTEM:LOCAL")
        elif control.find('remote') != -1:
            self.meas.write("SYSTEM:REMOTE")
    
    def loadImpedance(self, impedance):
        """ Sets the load impedance the device expects to be driving.
        This allows the output to be accurately set.
        """
        if type(impedance):
            self.meas.write("OUTP:LOAD %s" % impedance)
        elif impedance.find('inf') != -1:
            self.meas.write("OUTP:LOAD INF")
        elif impedance.find('min') != -1:
            self.meas.write("OUTP:LOAD MIN")
        elif impedance.find('max') != -1:
            self.meas.write("OUTP:LOAD MAX")
        else:
            print("ERROR: Invalid impedance parameter specified")
            sys.exit()
    
    def mode(self, mode):
        """ Selects the output mode
        Possible values are:
            sine     -> Sine wave
            square   -> Square wave
            ramp     -> Triangle/saw-tooth wave
            triangle -> Alias of ramp
            pulse    -> Pulse output
            noise    -> White noise
            dc       -> DC voltage
            user     -> Arbitrary waveforms
        """
        if mode.find('sin') != -1:
            self.meas.write("FUNC SIN")
        elif mode.find('squ') != -1:
            self.meas.write("FUNC SQU")
        elif mode.find('ramp') != -1:
            self.meas.write("FUNC RAMP")
        elif mode.find('tri') != -1:
            self.meas.write("FUNC RAMP")
        elif mode.find('puls') != -1:
            self.meas.write("FUNC PULS")
        elif mode.find('nois') != -1:
            self.meas.write("FUNC NOIS")
        elif mode.find('dc') != -1:
            self.meas.write("FUNC DC")
        elif mode.find('user') != -1:
            self.meas.write("FUNC USER")
        else:
            print("Invalid waveform mode specified")
            sys.exit()
            
    
    def frequency(self, num):
        """ Sets the output frequency to the given value
        """
        self.meas.write("FREQ %f" % num)
        #self._frequency = freq
        
    def voltage(self, amplitude=None):
        """ Sets the output voltage of the device.
        NOTE: The device expects to be driving into a 50 Ohm load so.
        If driving loads of higher impedance you will get more voltage.
        """
        if amplitude is not None:
            self.meas.write("VOLT %f" % amplitude)
            self.amplitude = amplitude
        return self.amplitude

    
    def burstCount(self, pulseCount):
        """ Sets the burst pulse count"""
        self.meas.write("BM:NCYC %f" % pulseCount)
    
    def burstOn(self, enableBurst):
        if enableBurst:
            self.meas.write("BM:STAT ON")
        else:
            self.meas.write("BM:STAT OFF")
    
    def extTig(self,enableTrig):
        if enableTrig:
            self.meas.write("TRIG:SOUR EXT")
        else:
            self.meas.write("TRIG:SOUR IN")

