"""
Project   : Scope
Filename  : Tek_MSO73340DX.py
Version   : v1.00, 09-28-2016
Author    : Mizuki Yurmoto
Contents  : This contains interface and command
Copyright : MegaChips Technology America Co.
            *** MegaChips Technology America STRICTLY CONFIDENTIAL ***
"""

# -- Includes ---------------------------------------------------------------------------------------

import time
import visa
from logging import getLogger

logger = getLogger()


# -- Class ------------------------------------------------------------------------------------------

class Tek_MSO73340DX(object):

    # constructer and destructer -------------------------------------------------------------------
    def __init__(self, Address="TCPIP0::Tekscope-581923::inst0::INSTR"):  # constructer
        ## Initialize connection with this machine
        self.h = None
        rm = visa.ResourceManager()
        self.h = rm.open_resource(Address, read_termination='\r')
        self.visa_TermChar = '\n'  # Termination characters for VISA Read/Write (default=\r\n)
        self._dpojet = []
        self.dpnum = 0
        ##write codes for initialization

    # Wrapping function of VISA commands -----------------------------------------------------------
    def VISA_write(self, cmd):
        self.h.write(cmd, termination=self.visa_TermChar)

    def VISA_read(self):
        return str(self.h.read(termination=self.visa_TermChar))

    def VISA_query(self, cmd):
        self.h.write(cmd, termination=self.visa_TermChar)
        time.sleep(0.010)
        return str(self.h.read(termination=self.visa_TermChar))

    # Instances ------------------------------------------------------------------------------------

    def get_idn(self):
        return self.VISA_query("*IDN?")

    def polling(self):
        time.sleep(0.1)
        while 1:
            ans = self.VISA_query('BUSY?')
            if (ans == '0') | (ans == ':BUSY 0'):
                break

    def Run(self):
        self.set_RunMode(0)
        self.VISA_write("ACQ:STATE 1")

    def Single(self):
        self.set_RunMode(1)
        self.VISA_write("ACQ:STATE 1")
        while self.VISA_query("ACQ:STATE?")[-1] != "0": time.sleep(0.1)

    def Stop(self):
        self.VISA_write("ACQ:STATE 0")

    def Clear(self):
        self.VISA_write('DIS:PERS:RESET')

    def set_RunMode(self, mode):  # 0:Endless 1: 1Seqence
        m = 'SEQ' if mode == 1 else 'RUNST'
        cmd = 'ACQ:STOPA ' + m
        self.VISA_write(cmd)

    def get_SampleRate(self):
        return self.VISA_query('HOR:DIG:SAMPLER?')

    def set_SampleRate(self, rate):
        self.VISA_write('HOR:MODE:SAMPLER %.2e' % rate)

    def set_RecordLength(self, length):
        self.VISA_write('HOR:MODE MANUAL')
        self.VISA_write('HOR:MODE:RECO %.2e' % length)

    def get_Val(self, meas):
        return float(self.VISA_query("MEASU:MEAS%d:VAL?" % meas))

    def set_Vscale(self, ch, sca):
        self.VISA_write("CH" + str(ch) + ":SCA %.2e" % sca)

    def set_Vposition(self, ch, pos):
        self.VISA_write(ch + ":POS %.2e" % pos)

    def set_Meas(self, x, typ, src):
        self.VISA_write("MEASU:MEAS%d:" % x)

    def run_Single(self):
        self.set_RunMode(1)
        self.VISA_write("ACQ:STATE 1")

    def rcl_Setup(self, filename):
        self.VISA_write('RECA:SETU "' + filename + '"')
        self.VISA_query("*OPC?")

    def get_AnalMean(self, meas):
        return float(self.VISA_query("DPOJET:MEAS%d:RESUL:CURR:MEAN?" % meas))

    def get_AnalP2P(self, meas):
        return float(self.VISA_query("DPOJET:MEAS%d:RESUL:CURR:PK2PK?" % meas))

    def SingleAnal(self):
        self.VISA_write("DPOJET:STATE SINGLE")
        while self.VISA_query("DPOJET:STATE?") != 'STOP':
            time.sleep(0.1)

    def Display(self, D, OnOff):
        self.VISA_write("SEL:" + D + " %d" % OnOff)

    def AutoVertical(self, ch):
        self.set_Vscale(ch, 0.5)
        self.VISA_write('MEASU:MEAS1:STATE ON')
        self.VISA_write('MEASU:MEAS1:TYP MAX')
        self.VISA_write('MEASU:MEAS1:SOU ' + ch)
        self.VISA_write('MEASU:MEAS2:TYP MINI')
        self.VISA_write('MEASU:MEAS2:SOU ' + ch)
        self.VISA_write('MEASU:MEAS1:STATE ON')
        self.VISA_write('MEASU:MEAS2:STATE ON')
        time.sleep(0.3)
        vmax = self.VISA_query('MEASU:MEAS1:VAL?')
        vmin = self.VISA_query('MEASU:MEAS2:VAL?')
        try:
            vmax = float(vmax)
            vmin = float(vmin)
        except:
            vmax = float(vmax.split(" ")[1])
            vmin = float(vmin.split(" ")[1])
        time.sleep(0.01)
        sca = (vmax - vmin) / 9.0
        pos = -(vmax + vmin) / (2 * sca)
        self.set_Vposition(ch, pos)
        self.set_Vscale(ch, sca)

        self.VISA_write('MEASU:MEAS1:STATE OFF')
        self.VISA_write('MEASU:MEAS2:STATE OFF')

        time.sleep(0.8)

    def DefaultSetup(self):
        self.VISA_write('FAC')
        time.sleep(3)
        self.polling()

    def save_waveform(self, port, filename):
        cmd = 'SAV:WAVE %s,' % port + '"%s"' % filename
        self.VISA_write(cmd)

    # ----- DPOJET functions -----#
    def dpjt_single(self):
        self.VISA_write('DPOJET:STATE SINGLE')
        while self.VISA_query('dpojet:state?') != 'STOP':
            time.sleep(0.1)

    def dpjt_clr(self):
        self.VISA_write('DPOJET:STATE CLEAR')
        while self.VISA_query('dpojet:state?') != 'STOP':
            time.sleep(0.1)

    def dpjt_autoset(self, src='VERT'):  # or HORIzontal,BOTH
        self.VISA_write('DPOJET:SOURCEA %s' % src)

    def dpjt_MeasVal(self, n):
        name = self.VISA_query('DPOJET:MEAS%d:NAME?' % n)
        src = self.VISA_query('DPOJET:MEAS%d:SOU1?' % n)
        if 'TIE' in name:
            val = float(self.VISA_query('DPOJET:MEAS%d:RESUL:CURR:PK2PK?' % n))
        else:
            val = float(self.VISA_query('DPOJET:MEAS%d:RESUL:CURR:MEAN?' % n))
        return [name, src, val]

    def dpjt_savePlot(self, n, sname):
        self.VISA_write(':DPOJET:EXPORT PLOT%d, "%s"' % (n, sname))
        time.sleep(0.5)

    # ----- SDLA functions ----#
    def open_SDLA(self):
        self.VISA_write('APPLICATION:ACTIVATE "Serial Data Link Analysis"')
        self.sdla_polling()

    def close_SDLA(self):
        self.VISA_write('VARIABLE:VALUE "sdla", "p:exit"')

    def sdla_analyze(self):
        self.VISA_write('VARIABLE:VALUE "sdla", "p:analyze"')
        self.sdla_polling()

    def sdla_apply(self):
        self.VISA_write('VARIABLE:VALUE "sdla", "p:apply"')
        self.sdla_polling()

    def sdla_recall(self, filename):
        self.VISA_write('VARIABLE:VALUE "sdla", "p:recall:%s"' % filename)
        self.sdla_polling()

    def sdla_setSrc(self, chname, src=1):  # chname = ch1|ch2|ch3|ch4|math1|ref1|ref2
        if src == 1:
            self.VISA_write('VARIABLE:VALUE "sdla", "p:source:%s"' % chname)
        elif src == 2:
            self.VISA_write('VARIABLE:VALUE "sdla", "p:source2:%s"' % chname)
        else:
            logger.warning('There is no "src%d" in the SDLA')
        self.sdla_polling()

    def sdla_polling(self):
        while self.VISA_query('VARIABLE:VALUE? "sdla"') != '"OK"':
            time.sleep(0.5)

    def math_Setdiff(self, n, ch1, ch2):
        cmd = "MATH%d:DEF " % n + "'Ch%d" % ch1 + "-Ch%d'" % ch2
        self.VISA_write(cmd)

    def init_4ch_f1f2(self):
        self.DefaultSetup()
        for i in range(1, 5):  # CH1 to CH4
            self.AutoVertical(i)
            self.VISA_write('SEL:CH%d ON' % i)
        # Horizontal Set up ##
        self.math_Setdiff(1, 1, 2)  # Math1 = CH1 -CH2
        self.math_Setdiff(2, 3, 4)  # Math1 = CH1 -CH2
        self.VISA_write('SEL:MATH1 ON')
        self.VISA_write('SEL:MATH2 ON')
        #            self.VISA_write('SEL:MATH1 OFF')
        #            self.VISA_write('SEL:MATH2 OFF')
        for i in range(1, 5):  # CH1 to CH4
            self.VISA_write('SEL:CH%d OFF' % i)
        self.set_SampleRate(100e9)  # Sampling Rate = 100G/s
        self.set_RecordLength(20e6)  # RecordLength  =  20M samples

    def meas_Vcp(self, count, MATH):
        self.DefaultSetup()
        if MATH == 1:
            self.VISA_write('TRIGGER:A:EDGE:SOURCE CH1')
        if MATH == 2:
            self.VISA_write('TRIGGER:A:EDGE:SOURCE CH3')
        for i in range(1, 5):  # CH1 to CH4
            self.AutoVertical(i)
            self.VISA_write('SEL:CH%d ON' % i)
        self.math_Setdiff(1, 1, 2)
        self.math_Setdiff(2, 3, 4)
        self.VISA_write('SEL:MATH1 ON')
        self.VISA_write('SEL:MATH2 ON')
        #            self.VISA_write('SEL:MATH1 OFF')
        #            self.VISA_write('SEL:MATH2 OFF')
        for i in range(1, 5):  # CH1 to CH4
            self.VISA_write('SEL:CH%d OFF' % i)
        self.VISA_write('HOR:MODE:SCA 1E-9')
        self.VISA_write('HIS:SOU MATH%d' % MATH)
        self.VISA_write('HIS:FUNC VERT')
        self.VISA_write('HIS:MOD VERT')
        self.polling()
        self.VISA_write('MEASU:MEAS1:SOU1 HIS')

        self.VISA_write('MEASU:MEAS1:TYP MEAN')
        self.VISA_write('MEASU:STAT:MOD ALL')
        time.sleep(1)
        self.VISA_write('HIS:BOX 2.4E-9, 0.6, 3E-9, -0.6')
        self.VISA_write('MEASU:MEAS1:STATE ON')
        time.sleep(count / 50)
        # print (self.VISA_query('MEASU:MEAS1:COUNT?'))
        VH = self.VISA_query('MEASU:MEAS1:MEAN?')
        #        VH = float(VH.split(" ")[1])
        VH = float(VH)
        # print VH

        self.VISA_write('HIS:BOX -800E-12, 0.6, -200E-12, -0.6')
        time.sleep(1)
        self.VISA_write('MEASU:MEAS1:STATE OFF')
        self.polling()
        self.VISA_write('MEASU:MEAS1:STATE ON')
        time.sleep(count / 50)
        # print (self.VISA_query('MEASU:MEAS1:COUNT?'))
        VL = self.VISA_query('MEASU:MEAS1:MEAN?')
        #        VL = float(VL.split(" ")[1])
        VL = float(VL)

        Vcp = VH - VL
        return Vcp

    def meas_VTNP(self, drate, count):
        self.DefaultSetup()
        self.VISA_write('TRIGGER:A:EDGE:SOURCE CH1')
        for i in range(1, 3):  # CH1 to CH2
            self.AutoVertical(i)
            self.VISA_write('SEL:CH%d ON' % i)
        self.math_Setdiff(1, 1, 2)
        self.VISA_write('SEL:MATH1 ON')
        #            self.VISA_write('SEL:MATH1 OFF')
        for i in range(1, 3):  # CH1 to CH2
            self.VISA_write('SEL:CH%d OFF' % i)
        self.VISA_write('HOR:MODE:SCA 1E-9')
        self.VISA_write('HIS:SOU MATH%d' % MATH)
        self.VISA_write('HIS:FUNC VERT')
        self.VISA_write('HIS:MOD VERT')
        self.polling()
        self.VISA_write('MEASU:MEAS1:SOU1 HIS')

        self.VISA_write('MEASU:MEAS1:TYP MEAN')
        self.VISA_write('MEASU:STAT:MOD ALL')
        time.sleep(1)
        self.VISA_write('HIS:BOX 2.4E-9, 0.6, 3E-9, -0.6')
        self.VISA_write('MEASU:MEAS1:STATE ON')
        time.sleep(count / 50)
        # print (self.VISA_query('MEASU:MEAS1:COUNT?'))
        VH = self.VISA_query('MEASU:MEAS1:MEAN?')
        #        VH = float(VH.split(" ")[1])
        VH = float(VH)
        # print VH

        self.VISA_write('HIS:BOX -800E-12, 0.6, -200E-12, -0.6')
        time.sleep(1)
        self.VISA_write('MEASU:MEAS1:STATE OFF')
        self.polling()
        self.VISA_write('MEASU:MEAS1:STATE ON')
        time.sleep(count / 50)
        # print (self.VISA_query('MEASU:MEAS1:COUNT?'))
        VL = self.VISA_query('MEASU:MEAS1:MEAN?')
        #        VL = float(VL.split(" ")[1])
        VL = float(VL)

        Vcp = VH - VL
        return Vcp


if __name__ == '__main__':
    handle_SC = 'TCPIP0::192.168.204.104::inst0::INSTR'
    scope = Tek_MSO73340DX(handle_SC)
    #    scope.DefaultSetup()
#    scope.math_Setdiff(1,1,3)
#    scope.Display('MATH1',1)
#    scope.Display('CH1',0)
#    scope.set_RecordLength(20e6)
#    scope.Stop()
#    scope.VISA_write('DIS:PERS:RESET')
#    scope.Single()
#    scope.save_waveform('MATH1','C:\\Share\\wfm\\test.wfm')
#    handle_SCtk = 'TCPIP0::192.168.201.74::inst0::INSTR'
##    import PIL
#    sc = Tek_MSO73340DX(Address = handle_SCtk)
#    sc.run_Single()
##    sc.dpjt_single()
##    sc.dpjt_savePlot(1,'C:/Share/a.png')
##    img = PIL.Image.open('Y:/a.png')
#
##    sc.dpjt_clr()
##    img.close()
#
#
#    handle_SCtk = 'TCPIP0::192.168.204.89::inst0::INSTR'
##    import PIL
#    sc = Tek_MSO73340DX(Address = handle_SCtk)
#    sc.run_Single()
##    sc.dpjt_single()
##    sc.dpjt_savePlot(1,'C:/Share/a.png')
##    img = PIL.Image.open('Y:/a.png')
#
##    sc.dpjt_clr()
##    img.close()


"""
1.0.0 : -- 20160928
"""