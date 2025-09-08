# -*- coding: utf-8 -*-
# ---------------------------------------------------------------------------------------------------
# Project   : Keysight_91304A
# Version   : v1.00($Rev: 14198 $), $Date: 2018-02-06 10:10:55 -0800 (Tue, 06 Feb 2018) $
#       SVN : $Id: Keysight_91304A.py 14198 2018-02-06 18:10:55Z clab $
# Author    : Mizuki Yurimoto
# Contents  : This contains interface and command
# Copyright : MegaChips Technology America Co.
#             *** MegaChips Technology America STRICTLY CONFIDENTIAL ***
# ---------------------------------------------------------------------------------------------------
# History
# v1.00 --- Created : Thu Dec 08 17:33:32 2016

import pyvisa as visa, time
from logging import getLogger

logger = getLogger()


class Keysight_91304A(object):
    def __init__(self, Address="TCPIP0::192.168.201.126::inst0::INSTR", timeout=180000):  # constructer
        ## Initialize connection with this machine
        self.h = None
        rm = visa.ResourceManager()
        self.h = rm.open_resource(Address, read_termination='\r')
        self.h.timeout = timeout  #
        #        self.h.timeout = 10000

        while True:
            err = int(self.readVISA(':SYST:ERR?'))
            if err == 0:
                break

    def writeVISA(self, cmd):
        self.h.write(cmd)
        self.chk_err()

    def chk_err(self, RaiseException=True):
        while True:
            try:
                err = int(self.h.query(':SYST:ERR?'))
            except:
                self.chk_err(False)
                err = 1

            message = 'Undifined Function' if err == -113 else \
                'Settings conflict' if err == -221 else \
                    'Data out of range' if err == -222 else \
                        'Query Interrupted' if err == -410 else \
                            'Invalid charactor' if err == -101 else \
                                'Query UNTERMINATED' if err == -420 else \
                                    'Warning :: %d' % err if err >= 1 else ''

            if err < 0:
                while True:
                    if self.readVISA(':SYST:ERR?') == '0': break

                if RaiseException:
                    raise Exception('%d:' % err + message)

            elif err > 0:
                logger.info(message)
            else:
                break

    def readVISA(self, cmd):
        err_cnt = 0
        while True:
            try:
                out = self.h.query(cmd)
                break
            except:
                logger.error("Err is occured !! again ...")
                err_cnt += 1
                if err_cnt == 5:
                    raise Exception
                time.sleep(1)
        return out

    def chkID(self):
        logger.info('[Get ID]')
        logger.info(self.readVISA('*IDN?'))

    def polling(self, monitor=False):
        while True:
            dat = self.readVISA(':ADER?')
            if monitor is True:
                print(dat)
            if dat == '+1':
                break
            else:
                time.sleep(0.2)

    def polling2(self, monitor=False):
        while True:
            dat = self.readVISA(':PDER?')
            if monitor is True:
                print(dat)
            if dat == '+1':
                break
            else:
                time.sleep(0.2)

    def setDefault(self):
        logger.info('Default ...')
        self.writeVISA('*RST')
        self.polling()

    def dispOnCH(self, ch):
        self.readVISA(':PDER?')
        self.writeVISA('CHAN%d ON' % ch)
        self.polling2()

    def dispOffCH(self, ch):
        self.writeVISA('CHAN%d OFF' % ch)

    def setTrig(self, ch):
        self.writeVISA(':TRIG:EDGE:SOUR CHAN%d' % ch)

    def getTrig(self):
        return self.readVISA(':TRIG:EDGE:SOUR?')

    def setTrigSwp(self, mode='Auto'):  # Auto or TRIGgered or SINGle
        self.writeVISA(':TRIG:SWE %s' % mode)

    def setTrigLevel(self, ampl):
        src = self.getTrig()
        self.writeVISA(':TRIG:LEV %s,%.3e' % (src, ampl))
        self.readVISA('*OPC?')

    def setTrigMode(self, mode='EDGE'):
        # EDGE,GLIT|PATT|STAT|DEL|TIMeout|TV|COMM|RUNT|SEQ|SHOL|TRAN|WIND|PWID|ADV|SBUS<N>
        self.writeVISA(':TRIG:MODE %s' % mode)

    def setTrigTimConf(self, ch=1, mode='HIGH', time=1e6):
        # mode : HIGH|LOW|UNCHanged
        self.writeVISA(':TRIG:TIM:SOUR CHAN%d' % ch)
        self.writeVISA(':TRIG:TIM:COND %s' % mode)
        self.writeVISA(':TRIG:TIM:TIME %.3e' % time)

    def setTrigPwidth(self, ch=1, mode='HIGHER', polarity='POS', tpoint='END', time=1e-6):
        mode_in = 'GTH' if mode == 'HIGHER' else 'LTH' if mode == 'LOWER' else 'GTH'
        tpoint_in = 'EPUL' if tpoint == 'END' else 'TIM' if tpoint == 'TIMEOUT' else 'EPUL'
        self.writeVISA(':TRIG:PWID:SOUR CHAN%d' % ch)
        self.writeVISA(':TRIG:PWID:DIR %s' % mode_in)
        self.writeVISA(':TRIG:PWID:POL %s' % polarity)
        self.writeVISA(':TRIG:PWID:TPO %s' % tpoint_in)
        self.writeVISA(':TRIG:PWID:WIDT %.3e' % time)

    def autoScale(self):
        self.writeVISA('AUT')

    def changeCHdiff(self, ch):
        self.writeVISA('CHAN%d:DIFF 1' % ch)

    def dispColorGrade(self, on=True):
        en = 1 if on is True else 0
        self.writeVISA(':DISP:CGR %d' % en)

    def setChDiffScaleAuto(self, ch, ON=True):
        on = 'ON' if ON is True else 'OFF'
        self.writeVISA(':CHAN%d:DISP:AUTO %s' % (ch, on))

    def setCtleSrc(self, ch):
        self.writeVISA(':SPR:CTL:SOUR CHAN%d' % ch)

    def setCtleACgain(self, val):
        self.writeVISA(':SPR:CTL:ACG %.3e' % val)

    def setCtleDCgain(self, val):
        self.writeVISA(':SPR:CTL:DCG %.2e' % val)

    def setCtleOfPoles(self, sel):
        # sel = 0 is help
        cmd = 'USB31' if sel == 1 \
            else 'POLE2' if sel == 2 \
            else 'POLE3' if sel == 3 else 0
        if cmd == 0:
            print('Input is %d ...' % sel)
            print('1:USB3.1\n2:POLE2\n3:POLE3')
        else:
            self.writeVISA('SPR:CTL:NUMP %s' % cmd)

    def setCtlePnFreq(self, n, freq):
        if n == 0:
            self.writeVISA('SPR:CTL:ZER %.3e' % freq)
        else:
            self.writeVISA('SPR:CTL:P%d %.3e' % (n, freq))

    def setCtleDataRate(self, rate):
        self.writeVISA(':SPR:CTL:RAT %.9e' % rate)

    def setCtleScaling(self, manualOn=False, sca=False, offset=False):
        # You can set any values to sca & offset
        # When you set 0 to setMAN (It means "AUTO")
        if manualOn is True:
            print('Manual Setting ...')
            self.readVISA(':PDER?')  # Processing Done Event
            self.writeVISA(':SPR:CTL:VERT MAN')
            self.polling2()
            if sca is not False:
                self.readVISA(':PDER?')
                self.writeVISA(':SPR:CTL:VERT:RANG %.5e' % (sca * 8))
                self.polling2()
            if offset is not False:
                self.writeVISA(':SPR:CTL:VERT:OFFS %.5e' % offset)
        else:
            print('Auto Setting ...')
            self.writeVISA(':SPR:CTL:VERT AUTO')

    def getCtleScaling(self):
        man = self.readVISA(':SPR:CTL:VERT?') == 'MAN'
        scale = float(self.readVISA(':SPR:CTL:VERT:RANG?'))
        offset = float(self.readVISA(':SPR:CTL:VERT:OFFS?'))
        return man, scale, offset

    def setCtleEn(self):
        self.writeVISA(':SPR:CTL:DISP 1')
        self.polling2()

    def setIsimStat(self, ch, stat='OFF'):
        statstr = ['OFF', 'PORT2', 'PORT4', 'PORT41']
        stat = statstr[stat] if type(stat) == int else stat
        self.writeVISA('CHAN%d:ISIM:STAT %s' % (ch, stat))

    def recallIsimFile(self, ch, filename):
        self.writeVISA('CHAN%d:ISIM:APPL "%s"' % (ch, filename))

    def openSetup(self, filename):
        self.readVISA(':PDER?')
        self.writeVISA(':DISK:LOAD "%s"' % filename)
        self.polling2()

    def run(self):
        self.writeVISA(':RUN')

    def stop(self):
        self.writeVISA(':STOP')

    def single(self):
        self.writeVISA(':STOP')
        self.readVISA(':ADER?')
        self.writeVISA(':SING')
        self.polling()

    def clearDis(self):
        self.readVISA(':PDER?')
        self.writeVISA(':CDIS')
        self.polling2()

    def setMemDepth(self, n):
        self.writeVISA(':ACQ:POIN %d' % n)
        self.polling2()

    def setSampleRate(self, n):
        self.writeVISA(':ACQ:SRAT %d' % n)
        self.polling2()

    def setChScale(self, ch, scale):
        self.writeVISA(':CHAN%d:SCAL %.3e' % (ch, scale))

    def single2(self, num=1e6):
        self.writeVISA(':RUN')
        while True:
            if self.getEyeNum() >= num:
                self.writeVISA(':STOP')
                break
            else:
                time.sleep(0.2)

    def setHistOn(self, on=True):
        cmd = 'WAV' if on == True else 'OFF'
        self.readVISA(':PDER?')
        self.writeVISA('HIST:MODE %s' % cmd)
        self.polling2()

    def setHistAxis(self, axis):
        # True  : horizontal
        # False : vertival
        cmd = 'HOR' if axis is True else 'VERT'
        self.readVISA(':PDER?')
        self.writeVISA('HIST:AXIS ' + cmd)
        self.polling2()

    def setHistDef(self):
        self.writeVISA(':HIST:WIND:DEF')

    def setHistSrc(self, ch=1):
        self.writeVISA(':HIST:WIND:SOUR CHAN%d' % ch)

    def setHistLimit(self, top=False, bottom=False, left=False, right=False):
        vals = [top, bottom, left, right]
        cmds = ['TLIM', 'BLIM', 'LLIM', 'RLIM']
        for i, j in zip(vals, cmds):
            if type(i) is int or type(i) is float:
                self.writeVISA(':HIST:WIND:%s %.2e' % (j, i))

    def getVpp(self, ch=1):
        ch = 'CHAN%d' % ch if type(ch) == int else ch
        return float(self.readVISA(':MEAS:VPP? %s' % ch))

    def getVamp(self, ch=1):
        ch = 'CHAN%d' % ch if type(ch) == int else ch
        return float(self.readVISA(':MEAS:VAMP? %s' % ch))

    def getVrms(self, ch=1, acdc='AC', cycdisp='CYCL'):
        ch = 'CHAN%d' % ch if type(ch) == int else ch
        return float(self.readVISA(':MEAS:VRMS? %s,%s,%s' % (cycdisp, acdc, ch)))

    def getFreq(self, ch=1):
        return float(self.readVISA(':MEAS:FREQ? CHAN%d' % ch))

    def getHistMin(self):
        return float(self.readVISA(':MEAS:HIST:MIN?'))

    def getHistMax(self):
        return float(self.readVISA(':MEAS:HIST:MAX?'))

    def getHistMean(self):
        return float(self.readVISA(':MEAS:HIST:MEAN?'))

    def getHistHits(self):
        return float(self.readVISA(':MEAS:HIST:HITS?'))

    def getHistMode(self):
        return float(self.readVISA(':MEAS:HIST:MODE?'))

    def getEyeWidth(self, alg='MEAS'):
        return float(self.readVISA(':MEAS:CGR:EWID? ' + alg))

    def getEyeHeight(self, alg='MEAS'):
        return float(self.readVISA(':MEAS:CGR:EHE? ' + alg))

    def getEye(self):
        eye_h = self.getEyeHeight()
        ehe_w = self.getEyeWidth()
        eye = [['EH', eye_h], ['EW', ehe_w]]
        self.writeVISA('')
        self.setHistOn(on=False)
        return eye

    def getDFEGain(self):
        return float(self.readVISA(':SPR:DFEQ:TAP:GAIN?'))

    def runDFEAuto(self):
        print('DFE adaptation ...')
        self.readVISA(':PDER?')
        self.writeVISA(':SPR:DFEQ:TAP:AUT')
        self.polling2()
        return self.getDFEValues()

    def setDFEStat(self, stat):
        cmd = 'ON' if stat == True else 'OFF'
        self.readVISA(':PDER?')
        self.writeVISA(':SPR:DFEQ:STAT %s' % cmd)
        self.polling2()

    def getDFEnTAPs(self):
        return int(self.readVISA(':SPR:DFEQ:NTAP?'))

    def getDFETapValue(self, n):
        return float(self.readVISA(':SPR:DFEQ:TAP? %d' % n))

    def getDFEUTarget(self):
        return float(self.readVISA(':SPR:DFEQ:TAP:UTAR?'))

    def getDFELTarget(self):
        return float(self.readVISA(':SPR:DFEQ:TAP:LTAR?'))

    def getDFEDelay(self):
        return float(self.readVISA(':SPR:DFEQ:TAP:DEL?'))

    def getDFEValues(self):
        ntap = self.getDFEnTAPs()
        utar = self.getDFEUTarget()
        ltar = self.getDFELTarget()
        delay = self.getDFEDelay()
        gain = self.getDFEGain()
        output = {'NTAP': ntap, 'UpperTarget': utar, 'LowerTarget': ltar, \
                  'Delay': delay, 'Gain': gain}
        for i in range(ntap):
            output['TAP%d' % (i + 1)] = self.getDFETapValue(i + 1)
        return output

    def saveImg(self, filename):
        self.readVISA(':PDER?')
        self.writeVISA(':DISK:SAVE:IMAG "%s"' % filename)
        self.polling2()

    def saveWfm(self, filename, fmt='BIN'):
        self.readVISA('PDER?')
        self.writeVISA(':DISK:SAVE:WAV ALL,"%s",%s' % (filename, fmt))
        self.polling2()

    def getEyeNum(self, src=''):
        return float(self.readVISA('MTES:FOLD:COUN:UI? ' + src))

    def setAutoThr(self, src='EQU'):
        self.writeVISA(':meas:thr:meth %s,HYST' % src)
        self.writeVISA(':MEASure:THResholds:GENauto ' + src)
        self.readVISA('*OPC?')

    def autoScaleV(self, ch):
        if type(ch) is int:
            self.writeVISA('AUT:VERT CHAN%d' % ch)
        elif type(ch) is str:
            self.writeVISA('AUT:VERT %s' % ch)
        self.polling2()

    def setWindowH(self, scale):
        self.writeVISA(':TIM:SCAL %e' % scale)

    def setPositionH(self, scale):
        self.writeVISA(':TIM:POS %e' % scale)

    def Mtest_En(self, enable=True):
        en = 1 if enable == True else 0
        self.writeVISA(':MTES:ENAB %d' % en)
        self.polling2()

    def Mtest_LoadFile(self, filename):
        self.writeVISA(':MTES:LOAD "%s"' % filename)
        self.polling2()

    def Mtest_start(self):
        self.writeVISA('MTES:STAR')

    def getMtestCnt(self):
        return float(self.readVISA('MTES:COUN:UI?'))

    def getMtestFUI(self):
        return float(self.readVISA('MTES:COUN:FUI?'))

    # Jitter/Noise Method
    def set_JitterNoiseEn(self, enable=True):
        en = 'ON' if enable is True else 'OFF'
        self.writeVISA(':MEAS:RJDJ:STATE %s' % en)
        time.sleep(5)
        self.readVISA('*OPC?')
        self.chk_err(False)

    def set_JiiterNoiseUnit(self, unit='UNIT'):  # or 'SEC'
        self.writeVISA(':MEAS:RJDJ:UNIT %s' % unit)

    def set_JitterNoiseBER(self, ber=9):  # 9 means BER:1e-9
        self.writeVISA(':MEAS:RJDJ:BER E%d' % ber)

    def set_JitterNoiseSrc(self, ch):
        if type(ch) == int:
            self.writeVISA(':MEAS:RJDJ:SOUR CHAN%d' % ch)
        else:
            self.writeVISA(':MEAS:RJDJ:SOUR %s' % ch)

    def set_JitterNoseRjMethod(self, method='spectral'):  # or 'tailfit'
        method_l = method.lower()
        meth = 'SPEC' if 'sp' in method_l else 'TAIL' if 'tail' in method_l else ''
        if meth != '':
            self.writeVISA(':MEAS:RJDJ:METH BOTH')
            self.writeVISA(':MEAS:RJDJ:REP %s' % meth)
        self.readVISA('*OPC?')

    def get_JitterNoiseTjRjDj(self):
        cnt = 100
        cc = 0
        while 1:
            dat = self.readVISA(':MEAS:RJDJ:TJRJDJ?').split(',')
            if int(dat[2]) + int(dat[5]) + int(dat[8]) == 0:
                return [[dat[x], float(dat[x + 1])] for x in range(0, 9, 3)]
            else:
                time.sleep(0.2)
                if cc > cnt:
                    return [[dat[x], float(dat[x + 1])] for x in range(0, 9, 3)]
                else:
                    cc += 1

    def get_JitterNoiseAll(self, t=False):
        jit = self.readVISA(':MEAS:RJDJ:ALL?').split(',')
        dat_list = [[jit[3 * x], float(jit[3 * x + 1])] for x in range(int(len(jit) / 3))]
        if t:
            dat_list = list(zip(*dat_list))

        return dat_list

    def add_funcFFT(self, ch=1, funcN=1):
        #        self.writeVISA("FUNC%d:MATL:OPER 'FFT Magnitude'"%funcN)
        self.writeVISA("FUNC%d:FFTM CHAN%d" % (funcN, ch))
        self.writeVISA('FUNC%d:DISP 1' % funcN)

    def add_measFFT(self, funcN=1, threshold=-40):
        self.writeVISA(':MEAS:FFT:FREQ FUNC1')
        self.writeVISA(':MEAS:FFT:MAGN FUNC1')
        self.writeVISA(':MEAS:FFT:THR %.2e' % threshold)

    def get_FFTmag(self, funcN=1):
        return float(self.readVISA(':MEAS:FFT:MAGN? FUNC%d' % funcN))

    def get_FFTfreq(self, funcN=1):
        return float(self.readVISA(':MEAS:FFT:FREQ? FUNC%d' % funcN))

    def get_measAll(self, col, *name):
        # col : 1=Current, 2=Min, 3=Max, 4=Mean, 5=Std Dev, 6=Count
        dat = self.readVISA(':MEAS:RES?').split(',')
        if name != ():
            dat = [[str(dat[x]), float(dat[x + col])] for jnam in name for x in range(len(dat)) if
                   dat[x].find(jnam) >= 0]
        return dat

    def get_measTIEpp(self):
        dat = self.readVISA(':MEAS:RES?').split(',')
        dat = [float(dat[x + 3]) - float(dat[x + 2]) for x in range(len(dat)) if dat[x].find('Data TIE') >= 0]
        return dat[0]

    def chg_FFTpeak1(self, n):
        self.writeVISA(':MEAS:FFT:PEAK1 %d' % n)

    def add_measVpp(self, ch=''):
        inp = 'CHAN%d' % ch if type(ch) == int else '%s' % ch
        self.writeVISA(':MEAS:VPP %s' % inp)

    def add_measVampl(self, ch=1):
        self.writeVISA(':MEAS:VAMP CHAN%d' % ch)

    def add_measFreq(self, ch=1):
        self.writeVISA(':MEAS:FREQ CHAN%d' % ch)

    def add_measDeEmph(self, ch=1):
        self.writeVISA(':MEAS:DEEM CHAN%d' % ch)

    def add_measPreShoot(self, ch=1):
        self.writeVISA(':MEAS:PRES CHAN%d' % ch)

    def add_measEyeHeight(self, stp=49, edp=51, option="MEAS"):
        """
        option =
            - MEAS = SMALLEST in WINDOW
            - EXTR = MEAN CONSIDERING DEVIATION
        """

        self.writeVISA(':MEAS:CGR:EWIN %d,%d' % (stp, edp))
        self.writeVISA(':MEAS:CGR:EHE %s' % option)
        self.readVISA('*OPC?')
        self.chk_err(False)

    def add_measTIE(self, ch=1):
        if type(ch) == int:
            self.writeVISA(":MEASure:TIEData2 CHAN%d,UNIT" % ch)
        else:
            self.writeVISA(":MEASure:TIEData2 %s,UNIT" % ch)
        self.readVISA('*OPC?')
        self.chk_err(False)

    def add_measDutyCyc(self, ch=1):
        if type(ch) == int:
            self.writeVISA(":MEASure:DUTYcycle CHAN%d" % ch)
        else:
            self.writeVISA(":MEASure:DUTYcycle %s" % ch)
        self.readVISA('*OPC?')
        self.chk_err(False)

    def add_measPeriod(self, ch=1):
        if type(ch) == int:
            self.writeVISA(":MEASure:PERiod CHAN%d" % ch)
        else:
            self.writeVISA(":MEASure:PERiod %s" % ch)
        self.readVISA('*OPC?')
        self.chk_err(False)

    def add_measFallTime(self, ch=1):
        if type(ch) == int:
            self.writeVISA(":MEASure:FALLtime CHAN%d" % ch)
        else:
            self.writeVISA(":MEASure:FALLtime %s" % ch)
        self.readVISA('*OPC?')
        self.chk_err(False)

    def launch_HdmiApp(self):
        self.writeVISA(":SYSTEM:LAUNCH 'N5399C/N5399D HDMI Test App'")

    # Probe Header Function
    def set_ProbeIntV(self, v, ch=1):
        self.writeVISA("CHAN%d:PROB:HEAD:VTER INT,%.2E" % (ch, v))


if __name__ == '__main__':
    handle_SC = 'TCPIP0::192.168.204.101::inst0::INSTR'
    scope = Keysight_91304A(Address=handle_SC)
