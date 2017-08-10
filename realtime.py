import random
import xlrd
from gurobipy import*
import time
import matplotlib.pyplot as plt
plt.ion()
drift = []
penalty = []

multi = [2.11,1.5,1.05,0.52]
VM = [0,1,10,100,1000,100000,1000000]

# -------------------------------------
def arrange(plT):
    j=0
    seq = [37,37,37,37,37]
    order = []
    for k in range(len(plT)):
        z = len(plT)
        seq[j%5] -= plT[z-k-1][j%5][1]
        order = [[seq[j%5],j%5]]+order
        j += 1
    return order
# -------------------------------------

for v in range(1):
    
    AM = [800,2000,4000,20000]
    YM = [200,400,800,4000]
    BM = [200,400,2000,4000]
    multiplier = 4.22/2
    cost = [0,0,0,0]
    
    for storage1 in range(1):
        random.seed(a=65)
        # reading the data
        microCost = [0 for t in range(48)]
        macroCost = [0 for t in range(48)]
        # reading file from excel 
        book = xlrd.open_workbook("C:\Users\Asus1\OneDrive\workspace\IEEETrans\demandData.xlsx")
        sheetD = book.sheet_by_index(0)
        sheetG = book.sheet_by_index(1)
        sheetP = book.sheet_by_index(2)
        # the time slot is considered to be of 1/2 hour
        # number of entities
        nw = 10
        ns = 50
        nc = 20
        nr = 500
        nf = 1
        nl = 50
    
        r11 = int(15*multiplier)
        r12 = int(12*multiplier)
        r21 = int(13*multiplier)
        r22 = int(120*multiplier)
        r31 = int(12*multiplier)
        r32 = int(120*multiplier)
        # variable names
        
        psTr = [[[] for t in range(48)]for r in range(nr)]
        psTc = [[[] for t in range(48)]for c in range(nc)]
        psTf = [[[] for t in range(48)]for f in range(nf)]
        
        tsTr = [[[] for t in range(48)]for r in range(nr)]
        tsTc = [[[] for t in range(48)]for c in range(nc)]
        tsTf = [[[] for t in range(48)]for f in range(nf)]
        
        plTf = [[] for k in range(nl)] # 50 jobs and 5 machines
        xplOf = [[[0 for t in range(48)] for j in range(5)] for k in range(nl)]
        xplEf = [[[0 for t in range(48)] for j in range(5)] for k in range(nl)]
        
        plT = [[] for t in range(48)]
        
        # setting up the loads
        # start time, end time, power, weight
        for r in range(nr):
            
            for rps in range(r12):
                t1 = int(random.uniform(12,39))
                t2 = t1 + int(random.uniform(3,8))
                l = random.uniform(17,23)
                psTr[r][t1].append([t1,t2,l])
        
            for rts in range(r11):
                t1 = int(random.uniform(12,42))
                t2 = t1 + int(random.uniform(1,4))
                l = random.uniform(4,8)
                tsTr[r][t1].append([t1,t2,l])
            
        for c in range(nc):
            
            for cps in range(r22):
                t1 = int(random.uniform(12,39))
                t2 = t1 + int(random.uniform(3,8))
                l = random.uniform(27,53)
                psTc[c][t1].append([t1,t2,l])
            
            for cts in range(r21):
                t1 = int(random.uniform(12,42))
                t2 = t1 + int(random.uniform(1,4))
                l = random.uniform(8,13)
                tsTc[c][t1].append([t1,t2,l])
                
        for f in range(nf):
            
            for fps in range(r32):
                t1 = int(random.uniform(12,39))
                t2 = t1 + int(random.uniform(3,6))
                l = random.uniform(47,63)
                psTf[f][t1].append([t1,t2,l])
              
            for fts in range(r31):
                t1 = int(random.uniform(12,42))
                t2 = t1 + int(random.uniform(1,4))
                l = random.uniform(10,14)
                tsTf[f][t1].append([t1,t2,l]) 
               
            # production line loads
            for k in range(nl):
                temp1 = []
                for j in range(5):
                    temp2 = []
                    temp2.append(int(random.uniform(20,40))) 
                    temp2.append(int(random.uniform(1,4)))
                    plTf[k].append(temp2)
                    temp1.append(temp2)
                plT[13].append(temp1) 
        
        # data input or generation mean
        generation_m = [[0 for c in range(2)] for t in range(48)]
        nsLoad_m = [[0 for c in range(3)] for t in range(48)]
        
        # generation total
        generation = [0 for t in range(48)]
        nsLoad = [0 for t in range(48)]
        
        # generation estimation
        generation_e = [0 for t in range(48)]
        nsLoad_e = [0 for t in range(48)]
        
        # storage
        Am = AM[storage1]
        Bm = BM[storage1]
        Ym = YM[storage1]
        
        # iteration over time
        for t in range(48):
            
            request = 0
            requested = 0
            incomplete = 0
            
            microCost[t] = sheetP.cell(t/2+1,0).value/1000
            macroCost[t] = sheetP.cell(t/2+1,1).value/1000
            
            generation_m[t][0] = sheetG.cell(t/2+1,1).value
            generation_m[t][1] = sheetG.cell(t/2+1,2).value
            nsLoad_m[t][0] = sheetD.cell(t/2+1,1).value/multiplier
            nsLoad_m[t][1] = sheetD.cell(t/2+1,2).value/multiplier
            nsLoad_m[t][2] = sheetD.cell(t/2+1,3).value/multiplier
            
            for w in range(nw):
                generation[t]+= random.normalvariate(generation_m[t][0],0.3*generation_m[t][0])
            for s in range(ns):
                generation[t]+= random.normalvariate(generation_m[t][1],0.3*generation_m[t][1])
            for c in range(nc):
                nsLoad[t] += max(10,random.normalvariate(nsLoad_m[t][0],0.3*nsLoad_m[t][0]))
            for r in range(nr):  
                nsLoad[t] += max(5,random.normalvariate(nsLoad_m[t][1],0.3*nsLoad_m[t][1]))
            for f in range(nf):
                nsLoad[t] += max(10,random.normalvariate(nsLoad_m[t][2],0.3*nsLoad_m[t][2]))
                
            generation_e[t] = nw*generation_m[t][0]+ns*generation_m[t][1]
            nsLoad_e[t] = nc*nsLoad_m[t][0]+nr*nsLoad_m[t][1]+nsLoad_m[t][2]
        
        # solving for the real time Model
        
        # real-time variables
        Qr = [0 for r in range(nr)]
        Qc = [0 for c in range(nc)]
        Qf = 0
        
        Trr = 50
        Tcc = 100
        Tff = 250
        costR = 0
        Ar = Am
        consumptionR = [nsLoad[t] for t in range(48)]
        
        V1 = 1
        V2 = 1
        
        for t in range(48):
            
            realtime = Model("RT")
            realtime.setParam('OutputFlag', False)
            objR = LinExpr()
            
            # solving the real time model
            xtsTr = [[] for r in range(nr)]
            xtsTc = [[] for c in range(nc)]
            xtsTf = [[] for f in range(nf)]
            
            xpsTr = [[] for r in range(nr)]
            xpsTc = [[] for c in range(nc)]
            xpsTf = [[] for f in range(nf)]
            
            xplT  = []
            
            B = realtime.addVar(0,Bm,0,GRB.CONTINUOUS,'discharge')
            X = realtime.addVar(0,generation[t],0,GRB.CONTINUOUS,'micro')
            G = realtime.addVar(0,GRB.INFINITY,0,GRB.CONTINUOUS,'macro')
            S = realtime.addVar(0,generation[t],0,GRB.CONTINUOUS,'macro')
            load = LinExpr()
            
            objR.addTerms(V1*macroCost[t]/2,G)
            objR.addTerms(V1*microCost[t]/2,X)
            #objR.addTerms(-V1*1.5*microCost[t]/2,S)
            
            objR.addTerms(0.000003,G)
            objR.addTerms(0.000002,X)
            objR.addTerms(0.000001,B)
            
            # load requested
            loadRr = [0 for r in range(nr)]
            loadRc = [0 for c in range(nc)]
            loadRf = [0 for f in range(nf)]
            
            # load served
            loadSr = [0 for r in range(nr)]
            loadSc = [0 for c in range(nc)]
            loadSf = [0 for f in range(nf)]
            
            for r in range(nr):
                for rts in range(len(tsTr[r][t])):
                    
                    xtsTr[r].append(realtime.addVar(0,1,0,GRB.BINARY,'xtsTr'))  # time shift able
                    objR.addTerms(-Qr[r]*V2*tsTr[r][t][rts][2]*(t+1-tsTr[r][t][rts][0]),xtsTr[r][rts])      # adding it to the objectives
                    loadRr[r] += tsTr[r][t][rts][2]*(t+1-tsTr[r][t][rts][0])
                    load.addTerms(tsTr[r][t][rts][2],xtsTr[r][rts])
                    requested += 1*(Qr[r]>0)
                    request += 1
                    
                for rps in range(len(psTr[r][t])):
                    
                    xpsTr[r].append(realtime.addVar(0,1,0,GRB.BINARY,'xpsTr'))
                    objR.addTerms(-Qr[r]*V2*(psTr[r][t][rps][2]-(t-psTr[r][t][rps][0])*psTr[r][t][rps][2]/(psTr[r][t][rps][1]-psTr[r][t][rps][0])),xpsTr[r][rps])
                    loadRr[r] += psTr[r][t][rps][2]-(t-psTr[r][t][rps][0])*psTr[r][t][rps][2]/(psTr[r][t][rps][1]-psTr[r][t][rps][0]) # load - load*t/time
                    load.addTerms(psTr[r][t][rps][2]/(psTr[r][t][rps][1]-t),xpsTr[r][rps])
                    requested += 1*(Qr[r]>0)
                    request += 1
            
            for c in range(nc):
                for cts in range(len(tsTc[c][t])):
                    
                    xtsTc[c].append(realtime.addVar(0,1,0,GRB.BINARY,'xtsTr'))
                    objR.addTerms(-Qc[c]*V2*tsTc[c][t][cts][2]*(t+1-tsTc[c][t][cts][0]),xtsTc[c][cts])
                    loadRc[c] += tsTc[c][t][cts][2]*(t+1-tsTc[c][t][cts][0])
                    load.addTerms(tsTc[c][t][cts][2],xtsTc[c][cts])
                    requested += 1*(Qc[c]>0)
                    request += 1
                    
                for cps in range(len(psTc[c][t])):
                    
                    xpsTc[c].append(realtime.addVar(0,1,0,GRB.BINARY,'xpsTr'))
                    objR.addTerms(-Qc[c]*V2*(psTc[c][t][cps][2]-(t-psTc[c][t][cps][0])*psTc[c][t][cps][2]/(psTc[c][t][cps][1]-psTc[c][t][cps][0])),xpsTc[c][cps])
                    loadRc[c] += psTc[c][t][cps][2]-(t-psTc[c][t][cps][0])*psTc[c][t][cps][2]/(psTc[c][t][cps][1]-psTc[c][t][cps][0])
                    load.addTerms(psTc[c][t][cps][2]/(psTc[c][t][cps][1]-t),xpsTc[c][cps])
                    requested += 1*(Qc[c]>0)
                    request += 1
                
            
            for f in range(nf):
                for fts in range(len(tsTf[f][t])):
                    
                    xtsTf[f].append(realtime.addVar(0,1,0,GRB.BINARY,'xtsTr'))
                    objR.addTerms(-Qf*V2*tsTf[f][t][fts][2]*(t+1-tsTf[f][t][fts][0]),xtsTf[f][fts])
                    loadRf[f] += tsTf[f][t][fts][2]*(t+1-tsTf[f][t][fts][0])
                    load.addTerms(tsTf[f][t][fts][2],xtsTf[f][fts])
                    requested += 1*(Qf>0)
                    request += 1
                    
                for fps in range(len(psTf[f][t])):
                    
                    xpsTf[f].append(realtime.addVar(0,1,0,GRB.BINARY,'xpsTr'))
                    objR.addTerms(-Qf*V2*(psTf[f][t][fps][2]-(t-psTf[f][t][fps][0])*psTf[f][t][fps][2]/(psTf[f][t][fps][1]-psTf[f][t][fps][0])),xpsTf[f][fps])
                    loadRf[f] += psTf[f][t][fps][2]-(t-psTf[f][t][fps][0])*psTf[f][t][fps][2]/(psTf[f][t][fps][1]-psTf[f][t][fps][0])
                    load.addTerms(psTf[f][t][fps][2]/(psTf[f][t][fps][1]-t),xpsTf[f][fps])
                    requested += 1*(Qf>0)
                    request += 1
            
            for k in range(len(plT[t])):
                temp1 = []
                for j in range(5):
                    temp1.append(realtime.addVar(0,1,0,GRB.BINARY,'xplT'))
                xplT.append(temp1)
            
            realtime.update()
            
            #job constraint jC
            
            for k in range(len(plT[t])):
                jC = LinExpr() # only in one machine
                requested += 1*(Qf>0)
                request += 1
                
                for j in range(5):
                    #objR.addTerms(-Qf*V2*(t-13)*plT[t][k][j][0],xplT[k][j])
                    objR.addTerms(-Qf*V2*plT[t][k][j][0],xplT[k][j])
                    load.addTerms(plT[t][k][j][0],xplT[k][j])
                    jC.addTerms(1,xplT[k][j])
                #loadRf[f] += (t-13)*(plT[t][k][0][0]+plT[t][k][1][0]+plT[t][k][2][0]+plT[t][k][3][0]+plT[t][k][4][0])/5
                loadRf[f] += (plT[t][k][0][0]+plT[t][k][1][0]+plT[t][k][2][0]+plT[t][k][3][0]+plT[t][k][4][0])/5
                realtime.addConstr(jC,GRB.LESS_EQUAL,1,'jC')
            
            # machine constraint
            for j in range(5):
                jM = LinExpr() 
                for k in range(len(plT[t])):
                    jM.addTerms(1,xplT[k][j])
                realtime.addConstr(jM,GRB.LESS_EQUAL,1,'jM')
                    
            realtime.setObjective(objR, GRB.MINIMIZE)
            
            # adding the constraints
            realtime.addConstr(B,GRB.LESS_EQUAL,Bm,'st')
            realtime.addConstr(B,GRB.LESS_EQUAL,Ar,'st')
            
            load.addTerms(-1,B)
            load.addTerms(-1,G)
            load.addTerms(-1,X)
            #load.addTerms(1,S)
            realtime.addConstr(load,GRB.EQUAL,-nsLoad[t],'s-d balance')
                
            realtime.optimize();
            
            # updating the values
            for r in range(nr):
                for rts in range(len(tsTr[r][t])):
                    
                    if xtsTr[r][rts].X < 0.1 and tsTr[r][t][rts][1]-1 > t:
                        tsTr[r][t+1].append(tsTr[r][t][rts])
                        incomplete += 1
                        
                    elif xtsTr[r][rts].X < 0.1 and tsTr[r][t][rts][1]-1 == t:
                        nsLoad[t+1] += tsTr[r][t][rts][2]
                        consumptionR[t+1] += tsTr[r][t][rts][2]
                        
                    elif xtsTr[r][rts].X > 0.1:
                        
                        consumptionR[t] += tsTr[r][t][rts][2]
                        loadSr[r] += tsTr[r][t][rts][2]*(t+1-tsTr[r][t][rts][0])
                        
                for rps in range(len(psTr[r][t])):
                    
                    if xpsTr[r][rps].X < 0.1 and psTr[r][t][rps][1]-1 > t:
                        psTr[r][t+1].append(psTr[r][t][rps])
                        incomplete += 1
                       
                    elif xpsTr[r][rps].X < 0.1 and psTr[r][t][rps][1]-1 == t:
                        nsLoad[t+1] += psTr[r][t][rps][2]
                        consumptionR[t+1] += psTr[r][t][rps][2]
                        
                    elif xpsTr[r][rps].X > 0.1:
                       
                        loadSr[r] += psTr[r][t][rps][2]-(t-psTr[r][t][rps][0])*psTr[r][t][rps][2]/(psTr[r][t][rps][1]-psTr[r][t][rps][0])
                        consumptionR[t] += psTr[r][t][rps][2]/(psTr[r][t][rps][1]-t)
                        
                        for t2 in range(t+1,psTr[r][t][rps][1]):
                            nsLoad[t2] += psTr[r][t][rps][2]/(psTr[r][t][rps][1]-t)
                            consumptionR[t2] += psTr[r][t][rps][2]/(psTr[r][t][rps][1]-t)
                      
            for c in range(nc):
               
                for cts in range(len(tsTc[c][t])):
                    
                    if xtsTc[c][cts].X < 0.1 and tsTc[c][t][cts][1]-1 > t:
                        tsTc[c][t+1].append(tsTc[c][t][cts])
                        incomplete += 1
                        
                    elif xtsTc[c][cts].X < 0.1 and tsTc[c][t][cts][1]-1 == t:
                        nsLoad[t+1] += tsTc[c][t][cts][2]
                        consumptionR[t+1]+= tsTc[c][t][cts][2]
                        
                    elif xtsTc[c][cts].X > 0.1:
                       
                        consumptionR[t]+= tsTc[c][t][cts][2]
                        loadSc[c] += tsTc[c][t][cts][2]*(t+1-tsTc[c][t][cts][0])
                
                for cps in range(len(psTc[c][t])):
                    if xpsTc[c][cps].X < 0.1 and psTc[c][t][cps][1]-1 > t:
                        psTc[c][t+1].append(psTc[c][t][cps])
                        incomplete += 1
                       
                    elif xpsTc[c][cps].X < 0.1 and psTc[c][t][cps][1]-1 == t:
                        nsLoad[t+1] += psTc[c][t][cps][2]
                        consumptionR[t+1] += psTc[c][t][cps][2]
                       
                    elif xpsTc[c][cps].X > 0.1:
                        
                        loadSc[c] += psTc[c][t][cps][2]-(t-psTc[c][t][cps][0])*psTc[c][t][cps][2]/(psTc[c][t][cps][1]-psTc[c][t][cps][0])
                        consumptionR[t] += psTc[c][t][cps][2]/(psTc[c][t][cps][1]-t)
                        for t2 in range(t+1,psTc[c][t][cps][1]):
                            nsLoad[t2] += psTc[c][t][cps][2]/(psTc[c][t][cps][1]-t)
                            consumptionR[t2] += psTc[c][t][cps][2]/(psTc[c][t][cps][1]-t)    
                            
            for f in range(nf):
                for fts in range(len(tsTf[f][t])):
                   
                    if xtsTf[f][fts].X < 0.1 and tsTf[f][t][fts][1]-1 > t:
                        tsTf[f][t+1].append(tsTf[f][t][fts])
                        incomplete += 1
                        
                    elif xtsTf[f][fts].X < 0.1 and tsTf[f][t][fts][1]-1 == t:
                        nsLoad[t+1] += tsTf[f][t][fts][2]
                        consumptionR[t+1]+= tsTf[f][t][fts][2]
                        
                    elif xtsTf[f][fts].X > 0.1:
                        
                        consumptionR[t]+= tsTf[f][t][fts][2]
                        loadSf[f] += tsTf[f][t][fts][2]*(t+1-tsTf[f][t][fts][0])
                
                for fps in range(len(psTf[f][t])):
                    if xpsTf[f][fps].X < 0.1 and psTf[f][t][fps][1]-1 > t:
                        psTf[f][t+1].append(psTf[f][t][fps])
                        incomplete += 1
                        
                    elif xpsTf[f][fps].X < 0.1 and psTf[f][t][fps][1]-1 == t:
                        nsLoad[t+1] += psTf[f][t][fps][2]
                        consumptionR[t+1] += psTf[f][t][fps][2]
                        
                    elif xpsTf[f][fps].X > 0.1:
                        
                        loadSf[f] += psTf[f][t][fps][2]-(t-psTf[f][t][fps][0])*psTf[f][t][fps][2]/(psTf[f][t][fps][1]-psTf[f][t][fps][0])
                        consumptionR[t] += psTf[f][t][fps][2]/(psTf[f][t][fps][1]-t)
                        for t2 in range(t+1,psTf[f][t][fps][1]):
                            nsLoad[t2] += psTf[f][t][fps][2]/(psTf[f][t][fps][1]-t)
                            consumptionR[t2] += psTf[f][t][fps][2]/(psTf[f][t][fps][1]-t)
                
                order = arrange(plT)
                
                for k in range(len(plT[t])):
                    tempflag = 0
                    for j in range(5):
                        if xplT[k][j].X > 0.1:
                            tempflag = 1
                            incomplete += 1
                            #loadSf[f] += plT[t][k][j][0]*(t-13)
                            loadSf[f] += plT[t][k][j][0]
                            consumptionR[t] += plT[t][k][j][0]
                            for t2 in range(t+1,plT[t][k][j][1]):
                                nsLoad[t2] += plT[t][k][j][0]
                                consumptionR[t2] += plT[t][k][j][0]
                    if tempflag == 0 and t < 37-3: # 3 is the number of slots
                        plT[t+1].append(plT[t][k])
                    
                    elif  tempflag == 0 and t == 37-3:
                        nsLoad[35] += (plT[t][k][0][0]+plT[t][k][1][0]+plT[t][k][2][0]+plT[t][k][3][0]+plT[t][k][4][0])/5
                        nsLoad[36] += (plT[t][k][0][0]+plT[t][k][1][0]+plT[t][k][2][0]+plT[t][k][3][0]+plT[t][k][4][0])/5
                        nsLoad[37] += (plT[t][k][0][0]+plT[t][k][1][0]+plT[t][k][2][0]+plT[t][k][3][0]+plT[t][k][4][0])/5
                        
                        consumptionR[35] += (plT[t][k][0][0]+plT[t][k][1][0]+plT[t][k][2][0]+plT[t][k][3][0]+plT[t][k][4][0])/5
                        consumptionR[36] += (plT[t][k][0][0]+plT[t][k][1][0]+plT[t][k][2][0]+plT[t][k][3][0]+plT[t][k][4][0])/5
                        consumptionR[37] += (plT[t][k][0][0]+plT[t][k][1][0]+plT[t][k][2][0]+plT[t][k][3][0]+plT[t][k][4][0])/5
            
            # updating the virtual queues
            for r in range(nr):
                Qr[r] = max(0,Qr[r]+loadRr[r]-loadSr[r]-Trr)

            for c in range(nc):
                Qc[c] = max(0,Qc[c]+loadRc[c]-loadSc[c]-Tcc)

            Qf = max(0,Qf+loadRf[0]-loadSf[0]-Tff)
            
            # updating storage
            ArT = max(0,generation[t]-consumptionR[t])
            Ar += min(ArT,Ym)
            Ar = max(0,Ar-B.X)
            Ar = min(Ar,Am)
            costR += G.X*macroCost[t]/2+X.X*microCost[t]/2#-S.X*microCost[t]*1.5/2
            drift.append(realtime.ObjVal+G.X*macroCost[t]/2+X.X*microCost[t]/2)
            penalty.append(G.X*macroCost[t]/2+X.X*microCost[t]/2)
            print drift[-1],penalty[-1]
            
            #print t,requested,request,incomplete,V1*G.X*macroCost[t]/2+V1*X.X*microCost[t]/2,realtime.ObjVal
            
        cost[storage1] += costR
        
    print cost[0],cost[1],cost[2],cost[3]

#plt.plot(drift)
plt.plot(penalty)
plt.show()