"""

Indr262 - Spring 2020 
Term Project

Team Members:
Melis Tuncay
Asrın Şevval Aktaş
Fırat Tamur

"""

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


def time_constraints(lp, variables, time, demands, products, markets):
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


def change_product(time, demands, profits):
    """
    Changes a product's time, demand and profit values. 

    Parameters:
    pd.DataFrame, pd.DataFrame, pd.DataFrame  


    Returns:
    pd.DataFrame, pd.DataFrame, pd.DataFrame   

    """

    old = input(f'Enter product name you want to change {list(demands.columns[1:])}: ').upper()

    if old not in list(time['Product'].values):
        print("Product not found!")
        return

    new = input("Enter a new product name: ").upper()
    
    demand = input("Enter demands for new product (Press enter to keep old demands): ")
    new_demands = demands.copy()

    if demand:
        demands = list(map(float, demand.split()))
        new_demands[old] = demand
   
    new_demands.rename(columns={old: new}, inplace=True)

    profit = input("Enter profits for new product (Press enter to keep old profits): ")
    new_profits = profits.copy()

    if profit:
        profit = list(map(float, profit.split())) 
        new_profits[old] = profit

    new_profits.rename(columns={old: new}, inplace=True)

    tm = input("Enter time for new product (Press enter to keep old time): ") 
    new_time = time.copy()

    if tm:
        tm = float(tm)
        new_time.loc[new_time['Product'] == old, 'Time(Mins)'] = tm
    
    new_time.loc[new_time['Product'] == old, 'Product'] = new

    return new_time, new_demands, new_profits


def introduce_product(time, demands, profits):
    """
    Add a new product to tables.

    Parameters:
    pd.DataFrame, pd.DataFrame, pd.DataFrame  


    Returns:
    pd.DataFrame, pd.DataFrame, pd.DataFrame   

    """

    new = input("Enter a new product name: ").upper()
    new_demands = demands.copy()

    demand = input("Enter demands for new product: ")

    if not demand:
        print("Demands must be provided.")

    demand = list(map(float, demand.split()))
    new_demands[new] = demand
   

    profit = input("Enter profits for new product: ")
    new_profits = profits.copy()

    if not profit:
        print("Profits must be provided.")
        return
    
    profit = list(map(float, profit.split()))
    new_profits[new] = profit
    
    tm = input("Enter time for new product: ") 
    new_time = time.copy()

    if not tm:
        print("Time must be provided.")
        return
        
    tm = float(tm)

    new_row = {'Product': new, 'Time(Mins)': tm}
    new_time = new_time.append(new_row, ignore_index=True)

    return new_time, new_demands, new_profits

    
def create_output(lp, demands):
    """
    Creates output excel file for given lp problem and saves it to Tables folder. 

    Parameters:
    pulp.LpProblem

    Returns:
    None

    """ 

    products = list(demands.columns)[1:]

    # read demands table
    markets = list(demands['Areas'].values)

    # rename sales areas to Area - i
    markets = ["-".join(markets[i].split()[1:]) for i in range(len(markets))]

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
    output.to_excel(f"../Tables/{lp.name}_output.xlsx", index=False)

    return output


def check_demands(lp, demands, verbose=1):
    """
    Compares sum of each product in lp output and demands and finds 
    satisfied and unsatisfied products.

    Parameters:
    pulp.LpProblem, pd.DataFrame

    Returns:
    list[str], list[str]

    """ 

    products = list(demands.columns)[1:]

    # read demands table
    markets = list(demands['Areas'].values)

    # rename sales areas to Area - i
    markets = ["-".join(markets[i].split()[1:]) for i in range(len(markets))]

    # take result of lp
    output = create_output(lp, demands)

    satisfied = list()
    unsatisfied = list()

    for product in products:
        if demands[product].sum() == output[product].sum():
            satisfied.append(product)
        else:
            unsatisfied.append(product)

    if verbose:
        print("Satisfied Products: ", satisfied)
        print("Unsatisfied Products: ", unsatisfied)

    return satisfied, unsatisfied


def solve_lp(name, time, demands, profits, verbose=1):
    """
    Solves and lp problem which has time, demands and profits dictionaries.
    Saves output of lp problem as name_output.xlsx

    Parameters:
    str, pulp.LpProblem, 3 * dict[str, dict[str, list[float]]]

    Returns:
    None

    """ 
    products = list(time['Product'].values)

    # read time table

    time_dict = {}

    for i in range(time.shape[0]):
        time_dict[products[i]] = time.loc[i, 'Time(Mins)']

    # read demands table
    markets = list(demands['Areas'].values)

    # rename sales areas to Area - i
    markets = ["-".join(markets[i].split()[1:]) for i in range(len(markets))]

    # read demand vaues

    demands_dict = {}

    for i in range(demands.shape[0]):
        demand = {}
        
        for product in products:
            demand[product] = demands.loc[i, product]

        demands_dict[markets[i]] = demand

    # read profit table
    profits_dict = {}

    for i in range(profits.shape[0]):
        profit = {}
        
        for product in products:
            profit[product] = profits.loc[i, product]

        profits_dict[markets[i]] = profit

    # create lp problem
    lp = pulp.LpProblem(name, sense=pulp.LpMaximize)

    # add objective 
    lp, variables = objective(lp, profits_dict)

    # add constraints
    lp = demand_constraits(lp, variables, demands_dict, products, markets)
    lp = time_constraints(lp, variables, time_dict, demands_dict, products, markets)

    # solve problem
    lp.solve()

    # find output demand values
    output = create_output(lp, demands)

    # find satisfied and unsatisfied products.
    satisfied, unsatisfied = check_demands(lp, demands, verbose=0)

    if verbose:
        # satisfied products
        print("\nSatisfied Products: ", satisfied)
        
        # print problem status
        print(f"\nSolution Status for {name}: {pulp.LpStatus[lp.status]}")

        # print objective value
        print(f"\nObjective Value for {name}: {pulp.value(lp.objective)}\n")

        # output demand values
        print(output)
        print() 

    # return lp problem.
    return lp, satisfied, unsatisfied


def write_lp_txt(name, lp, time, demands, profits, write='w'):
    """
    Writes result of a lp problem to txt file.

    """
    with open("part-e.txt", write) as file:
        
        file.write(f"Tables Used for problem:  {name}\n")
        file.write(name)
        file.write(time.to_csv(sep=' ', index=False, header=True))
        file.write("\n")
        file.write(demands.to_csv(sep=' ', index=False, header=True))
        file.write("\n")
        file.write(profits.to_csv(sep=' ', index=False, header=True))
        file.write("\n")  

        output = create_output(lp, demands)
        satisfied, unsatisfied = check_demands(lp, demands, verbose=0)

        file.write("Results:\n")

        file.write(f"Status: {pulp.LpStatus[lp.status]}\n")
        file.write(f"Objective: {pulp.value(lp.objective)}\n")
        file.write(f"Satisfied Products: {satisfied}\n\n")

        file.write("Output Table: \n")
    
        file.write(output.to_csv(sep=' ', index=False, header=True))
        file.write("\n")   



if __name__ == '__main__':
    
    # read excel files

    try:
        time = pd.read_excel('../Tables/time.xlsx')
        demands = pd.read_excel('../Tables/demands.xlsx')
        profits = pd.read_excel('../Tables/profits.xlsx')
    except:
        print("Tables not found. Wrong File Paths.")
    finally:
        exit()

    # if a product change of a new product comes set it True.
    sensivitiy_analysis = False

    time_new, demands_new, profits_new = time, demands, profits

    if 'y' == input("Do you want to change any product (y/n): "):
        sensivitiy_analysis = True
        time_new, demands_new, profits_new = change_product(time, demands, profits)

    if 'y' == input("Do you want to add any product (y/n): "):
        sensivitiy_analysis = True
        time_new, demands_new, profits_new = introduce_product(time_new, demands_new, profits_new)

    print("\n\nTables used for Initial Problem:")
    print(time)
    print(demands)
    print(profits)

    # if any product has changed or new product added.
    if sensivitiy_analysis:
        init_lp, sat, unsat = solve_lp("InitialProblem", time, demands, profits)
        write_lp_txt("InitialProblem", init_lp, time, demands, profits, 'w')
    
        print("\n\nTables used for Changed Problem:")
        print(time_new)
        print(demands_new)
        print(profits_new)
    
        changed_lp, sat, unsat = solve_lp("ChangedProblem", time_new, demands_new, profits_new)
        write_lp_txt("ChangedProblem", changed_lp, time_new, demands_new, profits_new, 'a')

    else:
        init_lp, sat, unsat = solve_lp("InitialProblem", time, demands, profits)
        write_lp_txt("InitialProblem", init_lp, time, demands, profits, 'w')
