"""

Indr262 - Spring 2020 
Term Project

Team Members:
Melis Tuncay
Asrın Şevval Aktaş
Fırat Tamur

"""

# Solution for part f in the project.

import pulp

# We found the unsatisfied products which are 'A', 'B', 'D'
# We built a lp problem which includes time values for 'A', 'B', 'D' 

lp = pulp.LpProblem("f)", sense=pulp.LpMaximize)

# Time variables for lp.
# We have tried a couple values and found that values after 0.62 for lower bound 
# makes lp problem infeasible.

t_a = pulp.LpVariable('t_a', lowBound=0.62, upBound=1)
t_b = pulp.LpVariable('t_b', lowBound=0.62, upBound=1.25)
t_d = pulp.LpVariable('t_d', lowBound=0.62, upBound=1.25)
lp += t_a + t_b + t_d

# We created a constraint from unsatisfied products time and remained time from 'C', 'E'

lp += 12000 * t_a + 6000 * t_b + 15000 * t_d <= 20450

# We used product old values as upper bound.

# lp += t_a <= 1
# lp += t_b <= 1.25
# lp += t_d <= 1.25

lp.solve()

print(f"\nSolution Status for: ", pulp.LpStatus[lp.status])

for variable in lp.variables():
    print("{} = {}".format(variable.name, variable.varValue))