# from __future__ import print_function
from ortools.sat.python import cp_model
import pandas as pd
import time
import numpy as np
import sys

from employee import Employee
from payroll import Payroll


class PartialSolutionPrinter(cp_model.CpSolverSolutionCallback):
    """Print intermediate solutions."""

    def __init__(self, x, num_employees, num_working_days, num_payrolls, sols):
        cp_model.CpSolverSolutionCallback.__init__(self)
        self._x = x
        self._num_employees = num_employees
        self._num_working_days = num_working_days
        self._num_payrolls = num_payrolls
        self._solutions = set(sols)
        self._solution_count = 0

    def on_solution_callback(self):
        self._solution_count += 1
        if self._solution_count in self._solutions:
            print('Solution %i' % self._solution_count)
            for d in range(self._num_working_days):
                print('Day %i' % d)
                for n in range(self._num_employees):
                    is_working = False
                    for s in range(self._num_payrolls):
                        if self.Value(self._x[d][s][n]):
                            is_working = True
                            print('  Employee %i works payroll %i' % (n, s))
                    # if not is_working:
                    #     print('  Nurse {} does not work'.format(n))
            print()

    def solution_count(self):
        return self._solution_count

def import_data():
    input_spreadsheet = pd.ExcelFile('Optimisation_Spreadsheet_v1.xlsx')
    input_employees = pd.read_excel(input_spreadsheet, 'Employees')
    input_payrolls = pd.read_excel(input_spreadsheet, 'Payrolls')

    # print(input_employees.loc[0,:])
    # print(input_payrolls)

    employees = []
    payrolls = []
    special_payrolls = []
    for e in range(len(input_employees)):
        employee = Employee(input_employees.loc[e, 'Employee'], input_employees.loc[e, 'Team'],
                            input_employees.loc[e, 'Role'], input_employees.loc[e, 'Technicality'],
                            input_employees.loc[e, 'Monthly Minutes'],
                            input_employees.loc[e, 'Availability in Minutes'])
        employees.append(employee)

    for p in range(246):
        payroll = Payroll(input_payrolls.loc[p, 'Payroll'],
                          input_payrolls.loc[p, 'Prev. Employee'],
                          input_payrolls.loc[p, 'Technicality'], input_payrolls.loc[p, 'Due date'].day,
                          input_payrolls.loc[p, 'Paydate'], input_payrolls.loc[p, 'Data sent date'],
                          input_payrolls.loc[p, 'Time in Minutes'])
        if payroll.get_processing_time() > 300:
            special_payrolls.append(payroll)
        else:
            payrolls.append(payroll)
        # print(payroll.get_due_date())
    # [print(payrolls[i].get_processing_time()) for i in range(len(payrolls))]
    # [print(special_payrolls[i].get_processing_time()) for i in range(len(special_payrolls))]
    return employees, payrolls


def main():
    start_time = time.time()
    employees, payrolls = import_data()

    model = cp_model.CpModel()

    working_days = []

    for i in range(1, 31):
        working_days.append(i)

    num_employees = len(employees)
    num_payrolls = len(payrolls)
    num_days = len(working_days)

    all_employees = range(num_employees)
    all_payrolls = range(num_payrolls)
    all_days = range(num_days)

    x = []
    for i in working_days:
        y = []
        for j in range(len(payrolls)):
            z = []
            for k in range(len(employees)):
                z.append(model.NewIntVar(0, 1, "x[%i,%i,%i]" % (i, j, k)))
            y.append(z)
        x.append(y)

    print("Initialising Constraints")

    # Every payroll assigned to at least one employee one day
    [model.Add(sum(x[i][j][k] for i in range(len(working_days)) for k in range(len(employees))) == 1)
     for j in range(len(payrolls))]

    # Constraint - sum of work across all working days <= max monthly hours for each employee
    [model.Add(sum(payrolls[j].get_processing_time() * x[i][j][k] for i in range(len(working_days)) for j in
                   range(len(payrolls))) <= employees[k].get_max_hours()) for k in range(len(employees))]

    # # Constraint = sum of work for each day <= 8 hours for each employee
    [model.Add(sum(payrolls[j].get_processing_time() * x[i][j][k] for j in
                   range(len(payrolls)) for k in range(len(employees))) <= 480) for i in range(len(working_days))]
    #
    # Constraint - payroll technicality <= employee technicality
    [model.Add(payrolls[j].get_technicality() * x[i][j][k] <= employees[k].get_technicality()) for i in
     range(len(working_days)) for j in range(len(payrolls)) for k in range(len(employees))]

    # Constraint - payroll assigned before it's due date
    [model.Add(i * x[i][j][k] <= payrolls[j].get_due_date()) for i in
     range(len(working_days)) for j in range(len(payrolls)) for k in range(len(employees))]

    print("Constraints intialised")

    # model.Minimize(sum(x[i][j][k] for i in range(len(working_days)) for k in range(len(employees)) for j in range(len(payrolls))))
    solver = cp_model.CpSolver()
    # solution_printer = PartialSolutionPrinter(x, num_employees, num_days, num_payrolls, range(10))
    # status = solver.SearchForAllSolutions(model, solution_printer)
    status = solver.Solve(model)

    print(status)

    print(time.time() - start_time)

    if status == cp_model.MODEL_SAT:
        count = 0
        print(solver.ObjectiveValue())
        for i in all_days:
            for j in all_payrolls:
                for k in all_employees:
                    if solver.Value(x[i][j][k]) == 1:
                        count += 1
                        print('Worker ', employees[k].get_name(), ' assigned to payroll ', payrolls[j].get_id(),
                              'on day ', i, ' due date', payrolls[j].get_due_date())
        print(count)

main()
