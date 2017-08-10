import random
import xlrd
from gurobipy import *
import timeit

multiply = [4.22,2.11,1.5,1.05,0.54]

requested = [0 for t in range(48)]
completed = [0 for t in range(48)]

# storage
AM = [800,2000,4000,20000]
BM = [200,400,800,4000]
YM = [200,400,800,4000]

for mul in range(1):
    random.seed(a=65+mul)
    start_time = timeit.default_timer()
    multiplier = 1.05
    storage1 = 0
    random.seed(a=65)
    Am = AM[storage1]
    Bm = BM[storage1]
    Ym = YM[storage1]
    
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
    
    totalRequest = [0 for c in range(nc)]
    
    optimal = Model("OP")
    optimal.setParam('OutputFlag', False)
    
    estimation = Model("OP")
    estimation.setParam('OutputFlag', False)
    
    GO = []
    XO = []
    AO = []
    BO = []
    YO = []
    SO = []
    
    GE = []
    XE = []
    AE = []
    BE = []
    YE = []
    SE = []
    
    r11 = int(15*multiplier) # time 
    r12 = int(12*multiplier) # power
    r21 = int(13*multiplier)
    r22 = int(120*multiplier)
    r31 = int(12*multiplier)
    r32 = int(120*multiplier)
    # variable names
    psOr = [[] for r in range(nr)]
    psOc = [[] for c in range(nc)]
    psOf = [[] for f in range(nf)]
    
    xpsOr = [[[0 for t in range(48)]for rts in range(r12)] for r in range(nr)] 
    xpsOc = [[[0 for t in range(48)]for cts in range(r22)] for c in range(nc)] 
    xpsOf = [[[0 for t in range(48)]for fts in range(r32)] for f in range(nf)] 
    
    xpsEr = [[[0 for t in range(48)]for rts in range(r12)] for r in range(nr)] 
    xpsEc = [[[0 for t in range(48)]for cts in range(r22)] for c in range(nc)] 
    xpsEf = [[[0 for t in range(48)]for fts in range(r32)] for f in range(nf)]
    
    tsOr = [[] for r in range(nr)]
    tsOc = [[] for c in range(nc)]
    tsOf = [[] for f in range(nf)]
    
    xtsOr = [[[0 for t in range(48)]for rts in range(r11)] for r in range(nr)] 
    xtsOc = [[[0 for t in range(48)]for cts in range(r21)] for c in range(nc)] 
    xtsOf = [[[0 for t in range(48)]for fts in range(r31)] for f in range(nf)] 
    
    xtsEr = [[[0 for t in range(48)]for rts in range(r11)] for r in range(nr)] 
    xtsEc = [[[0 for t in range(48)]for cts in range(r21)] for c in range(nc)] 
    xtsEf = [[[0 for t in range(48)]for fts in range(r31)] for f in range(nf)] 
    
    psTr = [[[] for t in range(48)]for r in range(nr)]
    psTc = [[[] for t in range(48)]for c in range(nc)]
    psTf = [[[] for t in range(48)]for f in range(nf)]
    
    tsTr = [[[] for t in range(48)]for r in range(nr)]
    tsTc = [[[] for t in range(48)]for c in range(nc)]
    tsTf = [[[] for t in range(48)]for f in range(nf)]
    
    plTf = [[] for k in range(nl)] # 50 jobs and 5 machines
    xplOf = [[[0 for t in range(48)] for j in range(5)] for k in range(nl)]
    xplEf = [[[0 for t in range(48)] for j in range(5)] for k in range(nl)]
    
    # setting up the loads
    # start time, end time, power, weight
    
    for r in range(nr):
        
        for rps in range(r12):
            t1 = int(random.uniform(12,39))
            t2 = t1 + int(random.uniform(3,8))
            l = random.uniform(17,23)
            psTr[r][t1].append([t1,t2,l])
            psOr[r].append([t1,t2,l])
            
            requested[t1] += 1
    
        for rts in range(r11):
            t1 = int(random.uniform(12,38))
            t2 = t1 + int(random.uniform(1,10))
            l = random.uniform(4,8)
            tsTr[r][t1].append([t1,t2,l])
            tsOr[r].append([t1,t2,l])
            
            requested[t1] += 1
        
    for c in range(nc):
        
        for cps in range(r22):
            t1 = int(random.uniform(12,39))
            t2 = t1 + int(random.uniform(3,8))
            l = random.uniform(27,53)
            psTc[c][t1].append([t1,t2,l])
            psOc[c].append([t1,t2,l])
            
            requested[t1] += 1
        
        for cts in range(r21):
            t1 = int(random.uniform(12,38))
            t2 = t1 + int(random.uniform(1,10))
            l = random.uniform(8,13)
            tsTc[c][t1].append([t1,t2,l])
            tsOc[c].append([t1,t2,l])
            
            requested[t1] += 1
            
    for f in range(nf):
        
        for fps in range(r32):
            t1 = int(random.uniform(12,39))
            t2 = t1 + int(random.uniform(3,6))
            l = random.uniform(47,63)
            psTf[f][t1].append([t1,t2,l])
            psOf[f].append([t1,t2,l])
            
            requested[t1] += 1
          
        for fts in range(r31):
            t1 = int(random.uniform(12,38))
            t2 = t1 + int(random.uniform(1,10))
            l = random.uniform(10,14)
            tsTf[f][t1].append([t1,t2,l]) 
            tsOf[f].append([t1,t2,l])
            
            requested[t1] += 1
           
        # production line loads
        for k in range(nl):
            requested[13] += 1
            for j in range(5):
                temp2 = []
                temp2.append(int(random.uniform(20,40))) 
                temp2.append(int(random.uniform(1,4)))
                plTf[k].append(temp2) 
    
    # data input or generation mean
    generation_m = [[0 for c in range(2)] for t in range(48)]
    nsLoad_m = [[0 for c in range(3)] for t in range(48)]
    
    # generation total
    generation = [0 for t in range(48)]
    nsLoad = [0 for t in range(48)]
    
    # generation estimation
    generation_e = [0 for t in range(48)]
    nsLoad_e = [0 for t in range(48)]
    
    
    objO = LinExpr()
    objE = LinExpr()
    
    # iteration over time
    for t in range(48):
        
        microCost[t] = sheetP.cell(t/2+1,0).value/1000
        macroCost[t] = sheetP.cell(t/2+1,1).value/1000
        
        generation_m[t][0] = sheetG.cell(t/2+1,1).value
        generation_m[t][1] = sheetG.cell(t/2+1,2).value
        nsLoad_m[t][0] = sheetD.cell(t/2+1,1).value/multiplier
        nsLoad_m[t][1] = sheetD.cell(t/2+1,2).value/multiplier
        nsLoad_m[t][2] = sheetD.cell(t/2+1,3).value/multiplier
        
        for w in range(nw):
            generation[t]+= random.normalvariate(generation_m[t][0],0.3*generation_m[t][0])#+(t%6)*1000#max(30-t,0)*200+max(t-30,0)*200#(t%6)*1000#
        for s in range(ns):
            generation[t]+= random.normalvariate(generation_m[t][1],0.3*generation_m[t][1])
        
        for r in range(nr):  
            d = max(5,random.normalvariate(nsLoad_m[t][1],0.3*nsLoad_m[t][1]))
            nsLoad[t] += d
            
        for c in range(nc):
            d = max(10,random.normalvariate(nsLoad_m[t][0],0.3*nsLoad_m[t][0]))
            nsLoad[t] += d
            
        for f in range(nf):
            d = max(10,random.normalvariate(nsLoad_m[t][2],0.3*nsLoad_m[t][2]))
            nsLoad[t] += d
            
        #print nsLoad[t]  
        generation_e[t] = nw*generation_m[t][0]+ns*generation_m[t][1]
        nsLoad_e[t] = nc*nsLoad_m[t][0]+nr*nsLoad_m[t][1]+nsLoad_m[t][2]
         
        # setting up of the variables
        for r in range(nr):
            for rps in range(r12):
                xpsOr[r][rps][t] = optimal.addVar(0,psOr[r][rps][2],0,GRB.CONTINUOUS,'xpsOr'+str(r)+str(rps)+str(t))
                xpsEr[r][rps][t] = estimation.addVar(0,psOr[r][rps][2],0,GRB.CONTINUOUS,'xpsOr'+str(r)+str(rps)+str(t))
                 
            for rts in range(r11):
                xtsOr[r][rts][t] = optimal.addVar(0,1,0,GRB.BINARY,'xtsOr'+str(r)+str(rts)+str(t))
                xtsEr[r][rts][t] = estimation.addVar(0,1,0,GRB.BINARY,'xtsOr'+str(r)+str(rts)+str(t))
         
        for c in range(nc):
            for cps in range(r22):
                xpsOc[c][cps][t] = optimal.addVar(0,psOc[c][cps][2],0,GRB.CONTINUOUS,'xpsOc'+str(c)+str(cps)+str(t))
                xpsEc[c][cps][t] = estimation.addVar(0,psOc[c][cps][2],0,GRB.CONTINUOUS,'xpsOc'+str(c)+str(cps)+str(t))
                  
            for cts in range(r21):
                xtsOc[c][cts][t] = optimal.addVar(0,1,0,GRB.BINARY,'xtsOc'+str(c)+str(cts)+str(t))
                xtsEc[c][cts][t] = estimation.addVar(0,1,0,GRB.BINARY,'xtsOc'+str(c)+str(cts)+str(t))
                 
        for f in range(nf):
            for fps in range(r32):
                xpsOf[f][fps][t] = optimal.addVar(0,psOf[f][fps][2],0,GRB.CONTINUOUS,'xpsOf'+str(f)+str(fps)+str(t))
                xpsEf[f][fps][t] = estimation.addVar(0,psOf[f][fps][2],0,GRB.CONTINUOUS,'xpsOf'+str(f)+str(fps)+str(t))
                 
            for fts in range(r31):
                xtsOf[f][fts][t] = optimal.addVar(0,1,0,GRB.BINARY,'xtsOf'+str(f)+str(fts)+str(t))  
                xtsEf[f][fts][t] = estimation.addVar(0,1,0,GRB.BINARY,'xtsOf'+str(f)+str(fts)+str(t))  
         
        for k in range(nl):
            for j in range(5):
                xplOf[k][j][t] = optimal.addVar(0,1,0,GRB.BINARY,'xplOf'+str(k)+str(j)+str(t))
                xplEf[k][j][t] = estimation.addVar(0,1,0,GRB.BINARY,'xplOf'+str(k)+str(j)+str(t))
         
        GO.append(optimal.addVar(0,GRB.INFINITY,0,GRB.CONTINUOUS,'GO'+str(t)))  
        XO.append(optimal.addVar(0,generation[t],0,GRB.CONTINUOUS,'XO'+str(t)))
        AO.append(optimal.addVar(0,Am,0,GRB.CONTINUOUS,'AO'+str(t)))
        BO.append(optimal.addVar(0,Bm,0,GRB.CONTINUOUS,'BO'+str(t)))
        YO.append(optimal.addVar(0,Ym,0,GRB.CONTINUOUS,'YO'+str(t))) 
        SO.append(optimal.addVar(0,generation[t],0,GRB.CONTINUOUS,'SO'+str(t))) 
         
        objO.addTerms(macroCost[t]/2,GO[t])
        objO.addTerms(microCost[t]/2,XO[t])
         
        GE.append(estimation.addVar(0,GRB.INFINITY,0,GRB.CONTINUOUS,'GO'+str(t)))  
        XE.append(estimation.addVar(0,generation_e[t],0,GRB.CONTINUOUS,'XO'+str(t)))
        AE.append(estimation.addVar(0,Am,0,GRB.CONTINUOUS,'AO'+str(t)))
        BE.append(estimation.addVar(0,Bm,0,GRB.CONTINUOUS,'BO'+str(t)))
        YE.append(estimation.addVar(0,Ym,0,GRB.CONTINUOUS,'YO'+str(t)))  
        SE.append(estimation.addVar(0,generation[t],0,GRB.CONTINUOUS,'SO'+str(t))) 
         
        objE.addTerms(macroCost[t]/2,GE[t])
        objE.addTerms(microCost[t]/2,XE[t])
         
    # --------- Solving the optimal, estimation, individualistic and stationary strategy    
    optimal.update()
    optimal.setObjective(objO, GRB.MINIMIZE)
     
    estimation.update()
    estimation.setObjective(objE, GRB.MINIMIZE)
     
    # setting the constraints
    # time shiftable constraint
    for r in range(nr):
        for rts in range(r11):
            ts = LinExpr()
            tse = LinExpr()
            for t in range(tsOr[r][rts][0],tsOr[r][rts][1]):
                ts.addTerms(1,xtsOr[r][rts][t])
                tse.addTerms(1,xtsEr[r][rts][t])
            optimal.addConstr(ts,GRB.EQUAL,1,"ts")
            estimation.addConstr(tse,GRB.EQUAL,1,"ts")
              
        for rps in range(r12):
            ps = LinExpr()
            pse = LinExpr()
            for t in range(psOr[r][rps][0],psOr[r][rps][1]):
                ps.addTerms(1,xpsOr[r][rps][t])
                pse.addTerms(1,xpsEr[r][rps][t])
            optimal.addConstr(ps,GRB.EQUAL,psOr[r][rps][2],"ps")  
            estimation.addConstr(pse,GRB.EQUAL,psOr[r][rps][2],"ps")      
      
    for c in range(nc):
        for cts in range(r21):
            ts = LinExpr()
            tse = LinExpr()
            for t in range(tsOc[c][cts][0],tsOc[c][cts][1]):
                ts.addTerms(1,xtsOc[c][cts][t])
                tse.addTerms(1,xtsEc[c][cts][t])
            optimal.addConstr(ts,GRB.EQUAL,1,"ts")
            estimation.addConstr(tse,GRB.EQUAL,1,"ts")
              
        for cps in range(r22):
            ps = LinExpr()
            pse = LinExpr()
            for t in range(psOc[c][cps][0],psOc[c][cps][1]):
                ps.addTerms(1,xpsOc[c][cps][t])
                pse.addTerms(1,xpsEc[c][cps][t])
            optimal.addConstr(ps,GRB.EQUAL,psOc[c][cps][2],"ps") 
            estimation.addConstr(pse,GRB.EQUAL,psOc[c][cps][2],"ps") 
      
    for f in range(nf):
        for fts in range(r31):
            ts = LinExpr()
            tse = LinExpr()
            for t in range(tsOf[f][fts][0],tsOf[f][fts][1]):
                ts.addTerms(1,xtsOf[f][fts][t])
                tse.addTerms(1,xtsEf[f][fts][t])
            optimal.addConstr(ts,GRB.EQUAL,1,"ts")
            estimation.addConstr(tse,GRB.EQUAL,1,"ts")
              
        for fps in range(r32):
            ps = LinExpr()
            pse = LinExpr()
            for t in range(psOf[f][fps][0],psOf[f][fps][1]):
                ps.addTerms(1,xpsOf[f][fps][t])
                pse.addTerms(1,xpsEf[f][fps][t])
            optimal.addConstr(ps,GRB.EQUAL,psOf[f][fps][2],"ps") 
            estimation.addConstr(pse,GRB.EQUAL,psOf[f][fps][2],"ps")
     
    for k in range(50):
        ts = LinExpr()
        tse = LinExpr()
        for j in range(5):
            for t in range(13,37):
                ts.addTerms(1,xplOf[k][j][t])
                tse.addTerms(1,xplEf[k][j][t])
                ts2 = LinExpr()
                ts2e = LinExpr()
                for t2 in range(1,plTf[k][j][1]):
                    ts2.addTerms(1,xplOf[k][j][t])
                    ts2.addTerms(-1,xplOf[k][j][t+t2]) # e at the end refers to the estimation model
                    ts2e.addTerms(1,xplEf[k][j][t])
                    ts2e.addTerms(-1,xplEf[k][j][t+t2])
                optimal.addConstr(ts2,GRB.LESS_EQUAL,0,'pl')
                estimation.addConstr(ts2e,GRB.LESS_EQUAL,0,'pl')
        optimal.addConstr(ts,GRB.EQUAL,1,'ts')
        estimation.addConstr(tse,GRB.EQUAL,1,'ts')
     
    for j in range(5):
        for t in range(48):
            ts = LinExpr()
            tse = LinExpr()
            for k in range(50):
                ts.addTerms(1,xplOf[k][f][t])
                tse.addTerms(1,xplEf[k][f][t])
            optimal.addConstr(ts,GRB.LESS_EQUAL,1,'ts')
            estimation.addConstr(tse,GRB.LESS_EQUAL,1,'ts')
             
    st = LinExpr()
    st.addTerms(1,AO[0])
    ste = LinExpr()
    ste.addTerms(1,AE[0])
     
    optimal.addConstr(st,GRB.EQUAL,Am,'st')
    estimation.addConstr(ste,GRB.EQUAL,Am,'st')
     
    for t in range(48):
        # battery evolution
        if t > 0:
            st = LinExpr();ste = LinExpr()
            st.addTerms(1,AO[t]);ste.addTerms(1,AE[t])
            st.addTerms(-1,AO[t-1]);ste.addTerms(-1,AE[t-1])
            st.addTerms(1,BO[t-1]);ste.addTerms(1,BE[t-1])
            st.addTerms(-1,YO[t-1]);ste.addTerms(-1,YE[t-1])
            optimal.addConstr(st,GRB.EQUAL,0,'st')
            estimation.addConstr(ste,GRB.EQUAL,0,'st')
             
        # limits on Y and B
        st = LinExpr()
        ste = LinExpr()
         
        st.addTerms(1,BO[t])
        ste.addTerms(1,BE[t])
         
        st.addTerms(-1,AO[t])
        ste.addTerms(-1,AE[t])
         
        optimal.addConstr(st,GRB.LESS_EQUAL,0,'st')
        estimation.addConstr(ste,GRB.LESS_EQUAL,0,'st')
         
        optimal.addConstr(YO[t],GRB.LESS_EQUAL,Ym,'st')
        optimal.addConstr(BO[t],GRB.LESS_EQUAL,Bm,'st')
         
        estimation.addConstr(YE[t],GRB.LESS_EQUAL,Ym,'st')
        estimation.addConstr(BE[t],GRB.LESS_EQUAL,Bm,'st')
         
        load = LinExpr()
        load.addTerms(1,XO[t])
        load.addTerms(1,GO[t])
        load.addTerms(1,BO[t])
        load.addTerms(-1,YO[t])
        load.addTerms(-1,SO[t])
        
        peak = LinExpr()
        peak.addTerms(1,XO[t])
        peak.addTerms(1,GO[t])
        #optimal.addConstr(peak,GRB.LESS_EQUAL,70000,'jj')
        
        loade = LinExpr()
        loade.addTerms(1,XE[t])
        loade.addTerms(1,GE[t])
        loade.addTerms(1,BE[t])
        loade.addTerms(-1,YE[t])
        loade.addTerms(-1,SE[t])
         
        for r in range(nr):
            for rts in range(r11):
                load.addTerms(-tsOr[r][rts][2],xtsOr[r][rts][t])
                loade.addTerms(-tsOr[r][rts][2],xtsEr[r][rts][t])
            for rps in range(r12):
                load.addTerms(-1,xpsOr[r][rps][t])
                loade.addTerms(-1,xpsEr[r][rps][t])
        for c in range(nc):
            for cts in range(r21):
                load.addTerms(-tsOc[c][cts][2],xtsOc[c][cts][t])
                loade.addTerms(-tsOc[c][cts][2],xtsEc[c][cts][t])
            for cps in range(r22):
                load.addTerms(-1,xpsOc[c][cps][t])
                loade.addTerms(-1,xpsEc[c][cps][t])
        for f in range(nf):
            for fts in range(r31):
                load.addTerms(-tsOf[f][fts][2],xtsOf[f][fts][t])
                loade.addTerms(-tsOf[f][fts][2],xtsEf[f][fts][t])
            for fps in range(r32):
                load.addTerms(-1,xpsOf[f][fps][t])
                loade.addTerms(-1,xpsEf[f][fps][t])
         
        for k in range(50):
            for j in range(5):
                load.addTerms(-plTf[k][j][0],xplOf[k][j][t])
                loade.addTerms(-plTf[k][j][0],xplEf[k][j][t])
                 
        optimal.addConstr(load,GRB.EQUAL,nsLoad[t],'load')
        estimation.addConstr(loade,GRB.EQUAL,nsLoad_e[t],'load')
         
    optimal.optimize();
    #print 'realCostO',optimal.ObjVal
    print optimal.ObjVal, timeit.default_timer()-start_time,'---------------------',multiplier
    estimation.optimize();
    #print 'realCostE',estimation.ObjVal
     
    # calculating actual cost in the case of estimation
    Ae = Am
    Ao = Am
    costE = 0
    costO = 0  # phargi just to compare the two methods, can be used for scaling
    consumptionO = [nsLoad[t] for t in range(48)]
    for t in range(48):
        consumptionE = nsLoad[t]
        for r in range(nr):
            for rts in range(r11):
                consumptionE += tsOr[r][rts][2]*xtsEr[r][rts][t].X
                consumptionO[t] += tsOr[r][rts][2]*xtsOr[r][rts][t].X
                
                completed[t] += 1*xtsOr[r][rts][t].X
                
            for rps in range(r12):
                consumptionE += xpsEr[r][rps][t].X
                consumptionO[t] += xpsOr[r][rps][t].X
                
                completed[t] += 1*(xpsOr[r][rps][t].X > 0)
         
        for c in range(nc):
            for cts in range(r21):
                consumptionE += tsOc[c][cts][2]*xtsEc[c][cts][t].X
                consumptionO[t] += tsOc[c][cts][2]*xtsOc[c][cts][t].X
                
                completed[t] += 1*xtsOc[c][cts][t].X
                
            for cps in range(r22):
                consumptionE += xpsEc[c][cps][t].X
                consumptionO[t] += xpsOc[c][cps][t].X
                
                completed[t] += 1*(xpsOc[c][cps][t].X > 0)
         
        for f in range(nf):
            for fts in range(r31):
                consumptionE += tsOf[f][fts][2]*xtsEf[f][fts][t].X
                consumptionO[t] += tsOf[f][fts][2]*xtsOf[f][fts][t].X
                
                completed[t] += 1*xtsOf[f][fts][t].X
                
            for fps in range(r32):
                consumptionE += xpsEf[f][fps][t].X
                consumptionO[t] += xpsOf[f][fps][t].X
                
                completed[t] += 1*(xpsOf[f][fps][t].X > 0)
         
        for k in range(50):
            for j in range(5):
                consumptionE += plTf[k][j][0]*xplEf[k][j][t].X
                consumptionO[t] += plTf[k][j][0]*xplOf[k][j][t].X
                
                completed[t] += 1*xplOf[k][j][t].X
         
        costE += microCost[t]*min(generation[t],consumptionE)+macroCost[t]*max(0,consumptionE-generation[t])
        costO += microCost[t]*min(generation[t],consumptionO[t])+macroCost[t]*max(0,consumptionO[t]-generation[t])
#         if mul == 0:
#             print consumptionO[t]/1000

import matplotlib.pyplot as plt

print sum(consumptionO),sum(nsLoad),sum(generation)
plt.figure(1)
# first sub plot
plt.subplot(221)
plt.plot(completed,label = 'completed')
plt.plot(requested,label = 'requested')
plt.legend()

# second sub plot
plt.subplot(222)
plt.plot(generation,label = 'generation')
plt.legend()

# third sub plot
plt.subplot(223)
plt.plot(nsLoad,label = 'non-shift')
plt.legend()

# fourth sub plot
plt.subplot(224)
plt.plot(consumptionO,label = 'consumption')
plt.legend()

plt.figure(2)
plt.subplot(211)
plt.plot(microCost,label='p1')
plt.plot(macroCost,label='p2')
plt.legend()

plt.subplot(212)
plt.plot([generation[t]-nsLoad[t] for t in range(48)],label='gap')
plt.legend()

plt.show()
# # calculating the cost in individualistic scheduling
# costI = 0
# Ai = Am
# consumptionI = [nsLoad[t] for t in range(48)]
# 
# costS = 0
# consumptionS = [nsLoad[t] for t in range(48)]
# 
# for r in range(nr):
#     for rts in range(r11):
#         minP = microCost[tsOr[r][rts][0]:tsOr[r][rts][1]].index(min(microCost[tsOr[r][rts][0]:tsOr[r][rts][1]]))
#         consumptionI[minP] += tsOr[r][rts][2]
#         consumptionS[tsOr[r][rts][0]] += tsOr[r][rts][2]
#     for rps in range(r12):
#         minP = microCost[psOr[r][rps][0]:psOr[r][rps][1]].index(min(microCost[psOr[r][rps][0]:psOr[r][rps][1]]))
#         consumptionI[minP] += psOr[r][rps][2]
#         for t2 in range(psOr[r][rps][0],psOr[r][rps][1]+1):
#             consumptionS[t2] += psOr[r][rps][2]/(psOr[r][rps][1]-psOr[r][rps][0])
#         
# for c in range(nc):
#     for cts in range(r21):
#         minP = microCost[tsOc[c][cts][0]:tsOc[c][cts][1]].index(min(microCost[tsOc[c][cts][0]:tsOc[c][cts][1]]))
#         consumptionI[minP] += tsOc[c][cts][2]
#         consumptionS[tsOc[c][cts][0]] += tsOc[c][cts][2]
#     for cps in range(r22):
#         minP = microCost[psOc[c][cps][0]:psOc[c][cps][1]].index(min(microCost[psOc[c][cps][0]:psOc[c][cps][1]]))
#         consumptionI[minP] += psOc[c][cps][2]
#         for t2 in range(psOc[c][cps][0],psOc[c][cps][1]+1):
#             consumptionS[t2] += psOc[c][cps][2]/(psOc[c][cps][1]-psOc[c][cps][0])
#         
# for f in range(nf):
#     for fts in range(r31):
#         minP = microCost[tsOf[f][fts][0]:tsOf[f][fts][1]].index(min(microCost[tsOf[f][fts][0]:tsOf[f][fts][1]]))
#         consumptionI[minP] += tsOf[f][fts][2]
#         consumptionS[tsOc[c][cts][0]] += tsOc[c][cts][2]
#     for cps in range(r32):
#         minP = microCost[psOf[f][fps][0]:psOf[f][fps][1]].index(min(microCost[psOf[f][fps][0]:psOf[f][fps][1]]))
#         consumptionI[minP] += psOf[f][fps][2]
#         for t2 in range(psOf[f][fps][0],psOf[f][fps][1]+1):
#             consumptionS[t2] += psOf[f][fps][2]/(psOf[f][fps][1]-psOf[f][fps][0])
# 
# for k in range(50):
#     j = int(random.uniform(0,5))
#     s1 = int(random.uniform(13,36))
#     s2 = s1+plTf[k][j][1]
#     for t2 in range(s1,s2+1):
#         consumptionI[t2]+= plTf[k][j][0] 
#         consumptionS[t2]+= plTf[k][j][0]
# 
# for t in range(48):
#     costI += microCost[t]*min(generation[t],consumptionI[t])+macroCost[t]*max(0,consumptionI[t]-generation[t])/2
#     costS += microCost[t]*min(generation[t],consumptionS[t])+macroCost[t]*max(0,consumptionS[t]-generation[t])/2
# 
# print costE/2
# print costS
# print costI

     
