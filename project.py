import pyomo
from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd



M=AbstractModel()

#commodities
M.c=Set()

#locations
M.l=Set()

#Params
M.Cos=Param(M.l,M.l)
M.Sup=Param(M.l,M.c)
M.Cap=Param(M.l)

#Dec Var
M.X=Var(M.c,M.l,M.l,within=NonNegativeReals)

def value_rule(model):
    return sum(model.Cos[i,j]*model.X[k,i,j] for k,i,j in model.c*model.l*model.l) 
#ObjFunc
M.value=Objective(rule=value_rule,sense=minimize)

def demand_rule(model,i,k):
    
    if model.Sup[i,k]<0 and i!=11 and i!=13: #not facility and demand node
        return sum(model.X[k,j,i] - model.X[k,i,j]  for j in model.l) == -model.Sup[i,k]
    elif model.Sup[i,k]>0 and i!=11 and i!=13: #not fc and supply node
        return sum(model.X[k,i,j] for j in model.l) <= model.Sup[i,k]
    elif i==11 or i==13:
        return sum(0.85*model.X[k,j,i]-model.X[k,i,j] for j in model.l) == 0
    else:
        return sum(model.X[k,j,i]-model.X[k,i,j] for j in model.l) ==0
def cap_rule(model,i):
    return sum(model.X[k,j,i] for k,j in model.c*model.l) <= model.Cap[i]

M.demand=Constraint(M.l,M.c,rule=demand_rule)
M.capacity=Constraint(M.l,rule=cap_rule)
data=DataPortal(model=M)
#read locations
data.load(filename="pythondata.xlsx", range="ltable", format='set', set='l')
#read commodities
data.load(filename="pythondata.xlsx", range="ctable", format='set', set='c')
#read demands
data.load(filename="pythondata.xlsx", range="Suptable", param='Sup',format="array")
#read capasities
data.load(filename="pythondata.xlsx", range="Captable",index='l', param='Cap')
#read costs
data.load(filename="pythondata.xlsx", range="Costabler", param='Cos',format="array")
instance = M.create_instance(data)

optimizer=SolverFactory("glpk")
optimizer.solve(instance)

var_val=[]
for v in instance.component_data_objects(Var):
    var_val.append(v.value)
nextValue=0;
thisdict={}
for k,i,j in instance.c*instance.l*instance.l:
    #print(k,i,j ," ",var_val[nextValue])
    index=str(k)+","+str(j)+","+str(i)
    thisdict.update({index: var_val[nextValue]})
    nextValue+=1;

df = pd.DataFrame(data=thisdict, index=[0])

df = (df.T)

df.to_excel('results.xlsx')







