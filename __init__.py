import numpy as np
import os
import matplotlib.pyplot as plt
import tkinter as tk
import pandas as pd
from tkinter import filedialog
from  matplotlib import ticker


class Tatayul:

    def __init__(self,titik):
        self.titik = titik,
        self.ExHy  = None,
        self.EyHx  = None,

        print('Program telah aktif!')
        print('Source code & documentation : https://github.com/SaefullAlam/Stratatools')
        try:
            import xlsxwriter
        except:
            import subprocess
            import sys
            self.install('xlsxwriter')

    @staticmethod
    def install(package):
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    

    def ekstrak_data_stratagem(self):
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        base_path = filedialog.askdirectory(parent=root)
        files = os.listdir(base_path)
        
        extr = {
                'ExHy_frekuensi' : [],
                'ExHy_periode'   : [],
                'ExHy_koherensi' : [],
                'ExHy_appres'    : [],
                'ExHy_fase'      : [],

                'EyHx_frekuensi' : [],
                'EyHx_periode'   : [],
                'EyHx_koherensi' : [],
                'EyHx_appres'    : [],
                'EyHx_fase'      : []
            }
        
        for item in files:
            file = os.path.join(base_path,item).replace("\\","/")
        
            with open(file) as ff:
                data = ff.read().splitlines()

            idx=0
            while idx<len(data):
                f = float(data[idx][:11].strip())
                p = 1/f
                ExHy_a = float(data[idx][22:33].strip())
                EyHx_a = float(data[idx][55:66].strip())

                if p>1e-6 and p<1e8 and ExHy_a!=0.0:
                    #ExHy frequency
                    extr['ExHy_frekuensi'].append(f)
                    #ExHy period
                    extr['ExHy_periode'].append(p)
                    #ExHy coherency
                    extr['ExHy_koherensi'].append(float(data[idx][11:22].strip()))
                    #ExHy scalar apparent resistivity
                    extr['ExHy_appres'].append(ExHy_a)
                    #ExHy scalar phase
                    extr['ExHy_fase'].append(float(data[idx][33:44].strip()))

                if p>1e-6 and p<1e8 and EyHx_a!=0.0:
                    #ExHy frequency
                    extr['EyHx_frekuensi'].append(f)
                    #ExHy period
                    extr['EyHx_periode'].append(p)
                    #EyHx coherency
                    extr['EyHx_koherensi'].append(float(data[idx][44:55].strip()))
                    #EyHx scalar apparent resistivity
                    extr['EyHx_appres'].append(EyHx_a)
                    #EyHx scalar phase
                    extr['EyHx_fase'].append(float(data[idx][66:77].strip()))

                if float(data[idx][:11].strip())==6310000.0:
                    break
                idx+=4
            print(f'Ekstrak file {item} selesai')
        ExHy_eks = pd.DataFrame({
                                'ExHy_frekuensi'    : extr['ExHy_frekuensi'],
                                'ExHy_periode'      : extr['ExHy_periode'],
                                'ExHy_periode_sqrt' : np.sqrt(extr['ExHy_periode']),
                                'ExHy_appres'       : extr['ExHy_appres'],
                                'ExHy_koherensi'    : extr['ExHy_koherensi'],
                                'ExHy_fase'         : extr['ExHy_fase']})
        EyHx_eks = pd.DataFrame({
                                'EyHx_frekuensi'    : extr['EyHx_frekuensi'],
                                'EyHx_periode'      : extr['EyHx_periode'],
                                'EyHx_periode_sqrt' : np.sqrt(extr['EyHx_periode']),
                                'EyHx_appres'       : extr['EyHx_appres'],
                                'EyHx_koherensi'    : extr['EyHx_koherensi'],
                                'EyHx_fase'         : extr['EyHx_fase']})
        setattr(self,'ExHy',ExHy_eks)
        setattr(self,'EyHx',EyHx_eks)
        return 'Selesai'
    
    def statistik(self,min_koherensi=0,max_koherensi=1):
        ExHy = self.ExHy
        EyHx = self.EyHx
        ExHy_filter = ExHy[(ExHy['ExHy_koherensi']>=min_koherensi) & (ExHy['ExHy_koherensi']<=max_koherensi)]
        EyHx_filter = EyHx[(EyHx['EyHx_koherensi']>=min_koherensi) & (EyHx['EyHx_koherensi']<=max_koherensi)]
        
        
        ExHy_02 = str(len(ExHy_filter[(ExHy_filter['ExHy_koherensi']>=0) & (ExHy_filter['ExHy_koherensi']<0.2)]))
        EyHx_02 = str(len(EyHx_filter[(EyHx_filter['EyHx_koherensi']>=0) & (EyHx_filter['EyHx_koherensi']<0.2)]))
        ExHy_24 = str(len(ExHy_filter[(ExHy_filter['ExHy_koherensi']>=0.2) & (ExHy_filter['ExHy_koherensi']<0.4)]))
        EyHx_24 = str(len(EyHx_filter[(EyHx_filter['EyHx_koherensi']>=0.2) & (EyHx_filter['EyHx_koherensi']<0.4)]))
        ExHy_46 = str(len(ExHy_filter[(ExHy_filter['ExHy_koherensi']>=0.4) & (ExHy_filter['ExHy_koherensi']<0.6)]))
        EyHx_46 = str(len(EyHx_filter[(EyHx_filter['EyHx_koherensi']>=0.4) & (EyHx_filter['EyHx_koherensi']<0.6)]))
        ExHy_68 = str(len(ExHy_filter[(ExHy_filter['ExHy_koherensi']>=0.6) & (ExHy_filter['ExHy_koherensi']<0.8)]))
        EyHx_68 = str(len(EyHx_filter[(EyHx_filter['EyHx_koherensi']>=0.6) & (EyHx_filter['EyHx_koherensi']<0.8)]))
        ExHy_810 = str(len(ExHy_filter[(ExHy_filter['ExHy_koherensi']>=0.8) & (ExHy_filter['ExHy_koherensi']<=1)]))
        EyHx_810 = str(len(EyHx_filter[(EyHx_filter['EyHx_koherensi']>=0.8) & (EyHx_filter['EyHx_koherensi']<=1)]))
        
        
        data ={
            'Parameter' : ['Jumlah data','Data 0-0.2','Data 0.2-0.4','Data 0.4-0.6','Data 0.6-0.8', 'Data 0.8-1',
                           'Min frekuensi','Maks frekuensi','Min app res','Maks app res ','Min koherensi','Maks koherensi'],
            'ExHy'      : [str(len(ExHy_filter)),ExHy_02,ExHy_24,ExHy_46,ExHy_68,ExHy_810,str(ExHy_filter['ExHy_frekuensi'].min()),str(ExHy_filter['ExHy_frekuensi'].max()),
                           str(ExHy_filter['ExHy_appres'].min()),str(ExHy_filter['ExHy_appres'].max()),
                           str(ExHy_filter['ExHy_koherensi'].min()),str(ExHy_filter['ExHy_koherensi'].max())],
            'EyHx'      : [str(len(EyHx_filter)),EyHx_02,EyHx_24,EyHx_46,EyHx_68,EyHx_810,str(EyHx_filter['EyHx_frekuensi'].min()),str(EyHx_filter['EyHx_frekuensi'].max()),
                            str(EyHx_filter['EyHx_appres'].min()),str(EyHx_filter['EyHx_appres'].max()),
                            str(EyHx_filter['EyHx_koherensi'].min()),str(EyHx_filter['EyHx_koherensi'].max())]
        }
        return pd.DataFrame(data)

    def plot1(self,min_koherensi=0,max_koherensi=1,min_appres=None,max_appres=None):
        ExHy = self.ExHy
        EyHx = self.EyHx
        ExHy_filter = ExHy[(ExHy['ExHy_koherensi']>=min_koherensi) & (ExHy['ExHy_koherensi']<=max_koherensi)]
        EyHx_filter = EyHx[(EyHx['EyHx_koherensi']>=min_koherensi) & (EyHx['EyHx_koherensi']<=max_koherensi)]

        if min_appres:
            ExHy_filter = ExHy_filter[ExHy_filter['ExHy_appres']>=min_appres]
            EyHx_filter = EyHx_filter[EyHx_filter['EyHx_appres']>=min_appres]
        if max_appres:
            ExHy_filter = ExHy_filter[ExHy_filter['ExHy_appres']<=max_appres]
            EyHx_filter = EyHx_filter[EyHx_filter['EyHx_appres']<=max_appres]

        min_periode    = min(ExHy_filter['ExHy_periode'])
        max_periode    = max(ExHy_filter['ExHy_periode'])
        min_appres  = min(ExHy_filter['ExHy_appres']) 
        max_appres  = max(ExHy_filter['ExHy_appres']) 
        
        if min(EyHx_filter['EyHx_periode']) < min_periode:
            min_periode = min(EyHx_filter['EyHx_periode'])
        if max(EyHx_filter['EyHx_periode']) > max_periode:
            max_periode = max(EyHx_filter['EyHx_periode'])
        if min(EyHx_filter['EyHx_appres']) < min_appres:
            min_appres = min(EyHx_filter['EyHx_appres'])
        if max(EyHx_filter['EyHx_appres']) > max_appres:
            max_appres = max(EyHx_filter['EyHx_appres'])
        
        xticks = 10**(np.arange(int(np.log10(min_periode))-2,int(np.log10(max_periode))+2,dtype='float'))
        yticks = 10**(np.arange(int(np.log10(min_appres))-2,int(np.log10(max_appres))+2,dtype='float'))
        
        fig,[Exy,Eyx] = plt.subplots(1,2,figsize=(10,9))

        Exy.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_appres'],linestyle='',marker='D',color='r',zorder=6)
        Exy.grid(zorder=1)
        Exy.set_ylabel('Rho app')
        Exy.set_xlabel('Periode')
        Exy.set_title('ExHy')
        Exy.set_xscale('log')
        Exy.set_yscale('log')
        Exy.set_xticks(xticks)
        Exy.set_yticks(yticks)
        Exy.set_aspect('equal','box')

        Eyx.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_appres'],linestyle='',marker='s',color='b',zorder=6)
        Eyx.grid(zorder=1)
        Eyx.set_ylabel('Rho app')
        Eyx.set_xlabel('Periode')
        Eyx.set_title('EyHx')
        Eyx.set_xscale('log')
        Eyx.set_yscale('log')
        Eyx.set_xticks(xticks)
        Eyx.set_yticks(yticks)
        Eyx.set_aspect('equal','box')
        plt.show()

    def plot2(self,min_koherensi=0,max_koherensi=1,min_appres=None,max_appres=None,komponen='EyHx'):
        from matplotlib.colors import ListedColormap
        
        

        
        ExHy = self.ExHy
        EyHx = self.EyHx
        ExHy_filter = ExHy[(ExHy['ExHy_koherensi']>=min_koherensi) & (ExHy['ExHy_koherensi']<=max_koherensi)]
        EyHx_filter = EyHx[(EyHx['EyHx_koherensi']>=min_koherensi) & (EyHx['EyHx_koherensi']<=max_koherensi)]

        if min_appres:
            ExHy_filter = ExHy_filter[ExHy_filter['ExHy_appres']>=min_appres]
            EyHx_filter = EyHx_filter[EyHx_filter['EyHx_appres']>=min_appres]
        if max_appres:
            ExHy_filter = ExHy_filter[ExHy_filter['ExHy_appres']<=max_appres]
            EyHx_filter = EyHx_filter[EyHx_filter['EyHx_appres']<=max_appres]

        min_periode    = min(ExHy_filter['ExHy_periode'])
        max_periode    = max(ExHy_filter['ExHy_periode'])
        min_appres  = min(ExHy_filter['ExHy_appres']) 
        max_appres  = max(ExHy_filter['ExHy_appres']) 
        
        if min(EyHx_filter['EyHx_periode']) < min_periode:
            min_periode = min(EyHx_filter['EyHx_periode'])
        if max(EyHx_filter['EyHx_periode']) > max_periode:
            max_periode = max(EyHx_filter['EyHx_periode'])
        if min(EyHx_filter['EyHx_appres']) < min_appres:
            min_appres = min(EyHx_filter['EyHx_appres'])
        if max(EyHx_filter['EyHx_appres']) > max_appres:
            max_appres = max(EyHx_filter['EyHx_appres'])
        
        xticks = 10**(np.arange(int(np.log10(min_periode))-2,int(np.log10(max_periode))+2,dtype='float'))
        yticks = 10**(np.arange(int(np.log10(min_appres))-2,int(np.log10(max_appres))+2,dtype='float'))
        
        if komponen=='EyHx':
            N = 1
            color = np.ones((100, 4))
            #0-1
            color[:1, 0] = np.linspace(111/255, 111/255, N)
            color[:1, 1] = np.linspace(252/255, 252/255, N)
            color[:1, 2] = np.linspace(231/255, 231/255, N)
            #1-2
            color[1:2, 0] = np.linspace(111/255, 111/255, N)
            color[1:2, 1] = np.linspace(252/255, 252/255, N)
            color[1:2, 2] = np.linspace(231/255, 231/255, N)
            #2-3
            color[2:3, 0] = np.linspace(111/255, 111/255, N)
            color[2:3, 1] = np.linspace(252/255, 252/255, N)
            color[2:3, 2] = np.linspace(231/255, 231/255, N)
            #3-4
            color[3:4, 0] = np.linspace(111/255, 111/255, N)
            color[3:4, 1] = np.linspace(252/255, 252/255, N)
            color[3:4, 2] = np.linspace(231/255, 231/255, N)
            #4-5
            color[4:5, 0] = np.linspace(111/255, 111/255, N)
            color[4:5, 1] = np.linspace(252/255, 252/255, N)
            color[4:5, 2] = np.linspace(231/255, 231/255, N)
            #5-6
            color[5:6, 0] = np.linspace(58/255, 58/255, N)
            color[5:6, 1] = np.linspace(116/255, 116/255, N)
            color[5:6, 2] = np.linspace(242/255, 242/255, N)
            #6-7
            color[6:7, 0] = np.linspace(58/255, 58/255, N)
            color[6:7, 1] = np.linspace(116/255,116/255, N)
            color[6:7, 2] = np.linspace(242/255, 242/255, N)
            #7-8
            color[7:8, 0] = np.linspace(7/255, 7/255, N)
            color[7:8, 1] = np.linspace(0/255, 0/255, N)
            color[7:8, 2] = np.linspace(204/255, 204/255, N)
            #8-9
            color[8:9, 0] = np.linspace(7/255, 7/255, N)
            color[8:9, 1] = np.linspace(0/255, 0/255, N)
            color[8:9, 2] = np.linspace(204/255, 204/255, N)
            #9-10
            color[9:10, 0] = np.linspace(5/255, 5/255, N)
            color[9:10, 1] = np.linspace(2/255, 2/255, N)
            color[9:10, 2] = np.linspace(89/255, 89/255, N)
            min_val = np.nanmin(EyHx_filter['EyHx_koherensi'].values)
            max_val = np.nanmax(EyHx_filter['EyHx_koherensi'].values)

            colss = color[int(min_koherensi*10):int(max_koherensi*10)]
            newcmp = ListedColormap(colss)

            fig,ax = plt.subplots(1,1,figsize=(10,10))
    #         ax.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_appres'],linestyle='',marker='D',zorder=6, label ="ExHy")
            plotscat = ax.scatter(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_appres'],marker='s',c=EyHx_filter['EyHx_koherensi'].values,cmap=newcmp,zorder=6, label = "EyHx")
            fig.colorbar(plotscat, ax=ax,label='Koherensi', shrink=0.75,location='right', pad=0.05,ticks=np.arange(0,1+0.1,0.1))
            ax.grid(zorder=1)
            ax.set_ylabel('Rho app')
            ax.set_xlabel('Periode')
            ax.set_title('EyHx')
            ax.set_xscale('log')
            ax.set_yscale('log')
            ax.set_xticks(xticks)
            ax.set_yticks(yticks)
            ax.legend()
            ax.set_aspect('equal','box')
            plt.show()

        if komponen=='ExHy':
            N = 1
            color = np.ones((100, 4))
            #0-1
            color[:1, 0] = np.linspace(252/255, 252/255, N)
            color[:1, 1] = np.linspace(141/255, 141/255, N)
            color[:1, 2] = np.linspace(141/255, 141/255, N)
            #1-2
            color[1:2, 0] = np.linspace(252/255, 252/255, N)
            color[1:2, 1] = np.linspace(141/255, 141/255, N)
            color[1:2, 2] = np.linspace(141/255, 141/255, N)
            #2-3
            color[2:3, 0] = np.linspace(252/255, 252/255, N)
            color[2:3, 1] = np.linspace(141/255, 141/255, N)
            color[2:3, 2] = np.linspace(141/255, 141/255, N)
            #3-4
            color[3:4, 0] = np.linspace(252/255, 252/255, N)
            color[3:4, 1] = np.linspace(141/255, 141/255, N)
            color[3:4, 2] = np.linspace(141/255, 141/255, N)
            #4-5
            color[4:5, 0] = np.linspace(252/255, 252/255, N)
            color[4:5, 1] = np.linspace(141/255, 141/255, N)
            color[4:5, 2] = np.linspace(141/255, 141/255, N)
            #5-6
            color[5:6, 0] = np.linspace(250/255, 250/255, N)
            color[5:6, 1] = np.linspace(82/255, 82/255, N)
            color[5:6, 2] = np.linspace(82/255, 82/255, N)
            #6-7
            color[6:7, 0] = np.linspace(250/255, 250/255, N)
            color[6:7, 1] = np.linspace(82/255, 82/255, N)
            color[6:7, 2] = np.linspace(82/255, 82/255, N)
            #7-8
            color[7:8, 0] = np.linspace(242/255, 242/255, N)
            color[7:8, 1] = np.linspace(27/255, 27/255, N)
            color[7:8, 2] = np.linspace(27/255, 27/255, N)
            #8-9
            color[8:9, 0] = np.linspace(242/255, 242/255, N)
            color[8:9, 1] = np.linspace(27/255, 27/255, N)
            color[8:9, 2] = np.linspace(27/255, 27/255, N)
            #9-10
            color[9:10, 0] = np.linspace(145/255, 145/255, N)
            color[9:10, 1] = np.linspace(3/255, 3/255, N)
            color[9:10, 2] = np.linspace(3/255, 3/255, N)
            min_val = np.nanmin(ExHy_filter['ExHy_koherensi'].values)
            max_val = np.nanmax(ExHy_filter['ExHy_koherensi'].values)

            colss = color[int(min_koherensi*10):int(max_koherensi*10)]
            newcmp = ListedColormap(colss)

            fig,ax = plt.subplots(1,1,figsize=(10,10))
    #         ax.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_appres'],linestyle='',marker='D',zorder=6, label ="ExHy")
            plotscat = ax.scatter(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_appres'],marker='D',c=ExHy_filter['ExHy_koherensi'].values,cmap=newcmp,zorder=6, label = "ExHy")
            fig.colorbar(plotscat, ax=ax,label='Koherensi', shrink=0.75,location='right', pad=0.05,ticks=np.arange(0,1+0.1,0.1))
            ax.grid(zorder=1)
            ax.set_ylabel('Rho app')
            ax.set_xlabel('Periode')
            ax.set_title('ExHy')
            ax.set_xscale('log')
            ax.set_yscale('log')
            ax.set_xticks(xticks)
            ax.set_yticks(yticks)
            ax.legend()
            ax.set_aspect('equal','box')
            plt.show()
            
    def plot3(self,min_koherensi=0,max_koherensi=1,min_appres=None,max_appres=None,komponen='both'):
        ExHy = self.ExHy
        EyHx = self.EyHx
        ExHy_filter = ExHy[(ExHy['ExHy_koherensi']>=min_koherensi) & (ExHy['ExHy_koherensi']<=max_koherensi)]
        EyHx_filter = EyHx[(EyHx['EyHx_koherensi']>=min_koherensi) & (EyHx['EyHx_koherensi']<=max_koherensi)]

        if min_appres:
            ExHy_filter = ExHy_filter[ExHy_filter['ExHy_appres']>=min_appres]
            EyHx_filter = EyHx_filter[EyHx_filter['EyHx_appres']>=min_appres]
        if max_appres:
            ExHy_filter = ExHy_filter[ExHy_filter['ExHy_appres']<=max_appres]
            EyHx_filter = EyHx_filter[EyHx_filter['EyHx_appres']<=max_appres]

        min_periode    = min(ExHy_filter['ExHy_periode'])
        max_periode    = max(ExHy_filter['ExHy_periode'])
        min_appres     = min(ExHy_filter['ExHy_appres']) 
        max_appres  = max(ExHy_filter['ExHy_appres']) 
        
        if min(EyHx_filter['EyHx_periode']) < min_periode:
            min_periode = min(EyHx_filter['EyHx_periode'])
        if max(EyHx_filter['EyHx_periode']) > max_periode:
            max_periode = max(EyHx_filter['EyHx_periode'])
        if min(EyHx_filter['EyHx_appres']) < min_appres:
            min_appres = min(EyHx_filter['EyHx_appres'])
        if max(EyHx_filter['EyHx_appres']) > max_appres:
            max_appres = max(EyHx_filter['EyHx_appres'])
        
        xticks = 10**(np.arange(int(np.log10(min_periode))-2,int(np.log10(max_periode))+2,dtype='float'))
        yticks = 10**(np.arange(int(np.log10(min_appres))-2,int(np.log10(max_appres))+2,dtype='float'))
        
        fig,[ax1,ax2,ax3] = plt.subplots(3,1,figsize=(10,9))
        
        if komponen=="both":
            ax1.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_appres'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6,label="ExHy")
            ax1.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_appres'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6,label="EyHx")
            ax1.grid(zorder=1)
            ax1.set_ylabel('Rho app')
            ax1.set_xlabel('Periode')
            ax1.set_title('Resistivitas Semu vs Periode')
            ax1.set_xscale('log')
            ax1.set_yscale('log')
            ax1.legend()

            ax2.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_koherensi'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax2.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_koherensi'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax2.grid(zorder=1)
            ax2.set_ylabel('Koherensi')
            ax2.set_title('Koherensi vs Periode')
            ax2.set_ylim([0,1.1])
            ax2.set_xscale('log')

            ax3.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_fase'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax3.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_fase'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax3.grid(zorder=1)
            ax3.set_xlabel('Periode (s)')
            ax3.set_ylabel('Fase')
            ax3.set_title('Fase vs Periode')
            ax3.set_xscale('log')
        elif komponen=="ExHy":
            ax1.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_appres'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6,label="ExHy")
#             ax1.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_appres'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6,label="EyHx")
            ax1.grid(zorder=1)
            ax1.set_ylabel('Rho app')
            ax1.set_xlabel('Periode')
            ax1.set_title('Resistivitas Semu vs Periode')
            ax1.set_xscale('log')
            ax1.set_yscale('log')
            ax1.legend()

            ax2.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_koherensi'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6)
#             ax2.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_koherensi'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax2.grid(zorder=1)
            ax2.set_ylabel('Koherensi')
            ax2.set_title('Koherensi vs Periode')
            ax2.set_ylim([0,1.1])
            ax2.set_xscale('log')

            ax3.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_fase'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6)
#             ax3.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_fase'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax3.grid(zorder=1)
            ax3.set_xlabel('Periode (s)')
            ax3.set_ylabel('Fase')
            ax3.set_title('Fase vs Periode')
            ax3.set_xscale('log')
        elif komponen=="EyHx":
#             ax1.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_appres'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6,label="ExHy")
            ax1.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_appres'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6,label="EyHx")
            ax1.grid(zorder=1)
            ax1.set_ylabel('Rho app')
            ax1.set_xlabel('Periode')
            ax1.set_title('Resistivitas Semu vs Periode')
            ax1.set_xscale('log')
            ax1.set_yscale('log')
            ax1.legend()

#             ax2.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_koherensi'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax2.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_koherensi'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax2.grid(zorder=1)
            ax2.set_ylabel('Koherensi')
            ax2.set_title('Koherensi vs Periode')
            ax2.set_ylim([0,1.1])
            ax2.set_xscale('log')

#             ax3.plot(ExHy_filter['ExHy_periode'],ExHy_filter['ExHy_fase'],linestyle='',marker='D',mec='r',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax3.plot(EyHx_filter['EyHx_periode'],EyHx_filter['EyHx_fase'],linestyle='',marker='s',mec='b',markerfacecolor="None",markeredgewidth=2,zorder=6)
            ax3.grid(zorder=1)
            ax3.set_xlabel('Periode (s)')
            ax3.set_ylabel('Fase')
            ax3.set_title('Fase vs Periode')
            ax3.set_xscale('log')
        else:
            raise Exception("Parameter komponen hanya menerima 'both','ExHy',dan 'EyHx'")
        
        fig.tight_layout()

        plt.show()
        
    def export_to_excel(self,min_koherensi=0,max_koherensi=1):
        ExHy = self.ExHy
        EyHx = self.EyHx
        ExHy_filter = ExHy[(ExHy['ExHy_koherensi']>=min_koherensi) & (ExHy['ExHy_koherensi']<=max_koherensi)]
        EyHx_filter = EyHx[(EyHx['EyHx_koherensi']>=min_koherensi) & (EyHx['EyHx_koherensi']<=max_koherensi)]
        writer = pd.ExcelWriter(self.titik[0]+'.xlsx', engine='xlsxwriter')
        ExHy_filter.to_excel(writer, sheet_name='ExHy',index=False)
        EyHx_filter.to_excel(writer, sheet_name='EyHx',index=False)
        writer.save()
        
    @staticmethod
    def plot_final(file,sheet='ExHy'):
        
        if sheet=='ExHy':
            try:
                Data = pd.read_excel(file,sheet_name='ExHy')
            except:
                import subprocess
                import sys
                self.install('openpyxl')
                try:
                    Data = pd.read_excel(file,sheet_name='EyHx')
                except:
                    print('Cek lagi data excelnya')

            fig,[ax1,ax2,ax3] = plt.subplots(3,1,figsize=(10,9))

            ax1.plot(Data['ExHy_periode'],Data['ExHy_appres'],linestyle='',marker='D',color='r', label='ExHy',zorder=6)
            ax1.grid(zorder=1)
            ax1.set_ylabel('Rho app')
            ax1.set_title('Resistivitas Semu vs Periode')
            ax1.set_xscale('log')
            ax1.set_yscale('log')
            ax1.legend(fontsize='large')

            ax2.plot(Data['ExHy_periode'],Data['ExHy_koherensi'],linestyle='',marker='D',color='r',zorder=6)
            ax2.grid(zorder=1)
            ax2.set_ylabel('Koherensi')
            ax2.set_title('Koherensi vs Periode')
            ax2.set_ylim([0,1.1])
            ax2.set_xscale('log')

            ax3.plot(Data['ExHy_periode'],Data['ExHy_fase'],linestyle='',marker='D',color='r',zorder=6)
            ax3.grid(zorder=1)
            ax3.set_xlabel('Periode (s)')
            ax3.set_ylabel('Fase')
            ax3.set_title('Fase vs Periode')
            ax3.set_xscale('log')
            fig.tight_layout()
            plt.show()
        
        if sheet=='EyHx':
            try:
                Data = pd.read_excel(file,sheet_name='EyHx')
            except:
                import subprocess
                import sys
                self.install('openpyxl')
                try:
                    Data = pd.read_excel(file,sheet_name='EyHx')
                except:
                    print('Cek lagi data excelnya')

            fig,[ax1,ax2,ax3] = plt.subplots(3,1,figsize=(10,9))

            ax1.plot(Data['EyHx_periode'],Data['EyHx_appres'],linestyle='',marker='s',color='b', label='EyHx',zorder=6)
            ax1.grid(zorder=1)
            ax1.set_ylabel('Rho app')
            ax1.set_title('Resistivitas Semu vs Periode')
            ax1.set_xscale('log')
            ax1.set_yscale('log')
            ax1.legend(fontsize='large')

            ax2.plot(Data['EyHx_periode'],Data['EyHx_koherensi'],linestyle='',marker='s',color='b',zorder=6)
            ax2.grid(zorder=1)
            ax2.set_ylabel('Koherensi')
            ax2.set_title('Koherensi vs Periode')
            ax2.set_ylim([0,1.1])
            ax2.set_xscale('log')

            ax3.plot(Data['EyHx_periode'],Data['EyHx_fase'],linestyle='',marker='s',color='b',zorder=6)
            ax3.grid(zorder=1)
            ax3.set_xlabel('Periode (s)')
            ax3.set_ylabel('Fase')
            ax3.set_title('Fase vs Periode')
            ax3.set_xscale('log')

            fig.tight_layout()