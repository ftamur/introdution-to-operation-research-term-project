# lp solver module
import pulp

# read excel file
import pandas as pd


# helper functions to build lp problem

def objective(lp, profits):
    """
    Builds objective function for given lp problem with profits dict. 

    Parameters:
    pulp.LpProblem, dict[str, dict[str, list[float]]] 


    Returns:
    pulp.LpProblem, dict[str, pulp.LpVariable]

    """

    obj = pulp.LpAffineExpression()
    variables = {}

    for market in profits:
        for product in profits[market]:
            name = market + product
            key = pulp.LpVariable(name, lowBound=0, cat=pulp.LpInteger)
            variables[name] = key
            value = profits[market][product]
            obj.addterm(key, value)

    lp += obj

    return lp, variables


def demand_constraits(lp, variables, demands, products, markets):
    """
    Adds constraint functions for given lp problem with demands dict. 

    Parameters:
    pulp.LpProblem, dict[str, pulp.LpVariable], 
    dict[str, dict[str, list[float]]], list[str], list[str] 


    Returns:
    pulp.LpProblem

    """

    for product in products:
        for market in markets:
            exp = pulp.LpAffineExpression(name='const_exp')
            rhs = demands[market][product]
            name = market + product
            key = variables[name]

            if demands[market][product] == 0:
                exp.addterm(key, 0)
            else:
                exp.addterm(key, 1)

            constraint = pulp.LpConstraint(e=exp, sense=pulp.LpConstraintLE, rhs=rhs)

            lp += constraint

    return lp


def time_constraints(lp, variables, time, products, markets):
    """
    Adds constraint function for given lp problem with time dict. 

    Parameters:
    pulp.LpProblem, dict[str, pulp.LpVariable], 
    dict[str, dict[str, list[float]]], list[str], list[str] 


    Returns:
    pulp.LpProblem

    """

    exp = pulp.LpAffineExpression(name='const_exp')
    rhs = 50000
    for product in products:
        for market in markets:
            name = market + product
            key = variables[name]
            if demands[market][product] == 0:
                exp.addterm(key, 0)
            else:
                exp.addterm(key, time[product])

    constraint = pulp.LpConstraint(e=exp, sense=pulp.LpConstraintLE, rhs=rhs)

    lp += constraint

    return lp


def create_output(lp, products, markets):
    """
    Creates output excel file for given lp problem and saves it to Tables folder. 

    Parameters:
    pulp.LpProblem

    Returns:
    None

    """ 

    variables = lp.variables()
    products = {product: list() for product in products}

    for variable in variables:
        products[variable.name[-1]].append(variable.varValue)

    output = pd.DataFrame(data=products)
    output['Areas'] = markets

    # rearrange columns
    cols = output.columns.tolist()

    cols = cols[-1:] + cols[:-1]
    output = output[cols]

    # write to excel file
    output.to_excel("../Tables/output.xlsx", index=False)


# read excel files

table_1 = pd.read_excel('../Tables/table_1.xlsx')
table_2 = pd.read_excel('../Tables/table_2.xlsx')
table_3 = pd.read_excel('../Tables/table_3.xlsx')

products = list(table_1['Product'].values)

# read time table - table_1

time = {}

for i in range(table_1.shape[0]):
    time[products[i]] = table_1.loc[i, 'Time(Mins)']

# read demands table - table_2

markets = list(table_2['Areas'].values)

# rename sales areas to Area - i
markets = ["-".join(markets[i].split()[1:]) for i in range(len(markets))]

# read demand vaues

demands = {}

for i in range(table_2.shape[0]):
    demand = {}
    
    for product in products:
        demand[product] = table_2.loc[i, product]

    demands[markets[i]] = demand

# read profit table - table_3

profits = {}

for i in range(table_3.shape[0]):
    profit = {}
    
    for product in products:
        profit[product] = table_3.loc[i, product]

    profits[markets[i]] = profit


# products = ['A', 'B', 'C', 'D', 'E']
# markets = ['1', '2', '3', '4']

# time = {'A': 1, 'B': 1.25, 'C': 0.75, 'D': 1.25, 'E': 1.2} 

# demands = {'1': {'A': 6000, 'B': 3000, 'C': 6000, 'D': 5000, 'E': 0},
#            '2': {'A': 2000, 'B': 0,    'C': 6000, 'D': 5000, 'E': 9000},
#            '3': {'A': 4000, 'B': 1000, 'C': 0,    'D': 3000, 'E': 2000},
#            '4': {'A': 0,    'B': 2000, 'C': 5000, 'D': 2000, 'E': 3000}}

# profits = {'1': {'A': 4.80, 'B': 4.75, 'C': 3.75, 'D': 5.25, 'E': 0},
#            '2': {'A': 3.75, 'B': 0,    'C': 3.30, 'D': 5.0,  'E': 5.04},
#            '3': {'A': 4.0,  'B': 5.0,  'C': 0,    'D': 5.15, 'E': 5.25},
#            '4': {'A': 0,    'B': 4.50, 'C': 3.50, 'D': 5.05, 'E': 5.15}}

# create lp problem
lp = pulp.LpProblem('Term-Project', sense=pulp.LpMaximize)

 
# add objective 
lp, variables = objective(lp, profits)

# add constraints
lp = demand_constraits(lp, variables, demands, products, markets)
lp = time_constraints(lp, variables, time, products, markets)

# solve problem
lp.solve()

# print problem status
print(pulp.LpStatus[lp.status])

# print variable values
for variable in lp.variables():
    print("{} = {}".format(variable.name, variable.varValue))

# print objective value
print(pulp.value(lp.objective))

create_output(lp, products, markets)



