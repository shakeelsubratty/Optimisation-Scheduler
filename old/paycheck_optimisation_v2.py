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

    def __init__(self, x, num_employees, num_payrolls, sols):
        cp_model.CpSolverSolutionCallback.__init__(self)
        self._x = x
        self._num_employees = num_employees
        self._num_payrolls = num_payrolls
        self._solutions = set(sols)
        self._solution_count = 0

    def on_solution_callback(self):
        self._solution_count += 1
        if self._solution_count in self._solutions:
            print('Solution %i' % self._solution_count)
            for e in range(self._num_employees):
                is_working = False
                for p in range(self._num_payrolls):
                    if self.Value(self._x[p][e]):
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

    employees = []
    payrolls = []
    for e in range(len(input_employees)):
        employee = Employee(input_employees.loc[e, 'Employee'], input_employees.loc[e, 'Team'],
                            input_employees.loc[e, 'Role'], input_employees.loc[e, 'Technicality'],
                            input_employees.loc[e, 'Monthly Minutes'],
                            input_employees.loc[e, 'Availability in Minutes'])
        employees.append(employee)

    for p in range(len(input_payrolls)):
        payroll = Payroll(input_payrolls.loc[p, 'Payroll'],
                          input_payrolls.loc[p, 'Prev. Employee'],
                          input_payrolls.loc[p, 'Technicality'], input_payrolls.loc[p, 'Due date'].day,
                          input_payrolls.loc[p, 'Paydate'], input_payrolls.loc[p, 'Data sent date'],
                          input_payrolls.loc[p, 'Time in Minutes'],
                          input_payrolls.loc[p, 'Do not reallocate Flag'] == 'Y')
        payrolls.append(payroll)
    return employees, payrolls


def main():

    toggle_do_not_reallocate = False

    start_time = time.time()
    employees, payrolls = import_data()

    model = cp_model.CpModel()

    num_employees = len(employees)
    num_payrolls = len(payrolls)

    all_employees = range(num_employees)
    all_payrolls = range(num_payrolls)

    x = []
    for p in all_payrolls:
        y = []
        if toggle_do_not_reallocate:
            if payrolls[p].get_do_not_reallocate():
                prev_emp_str = payrolls[p].get_previous_employee()
                for e in all_employees:
                    if employees[e].get_name() == prev_emp_str:
                        y.append(model.NewIntVar(1, 1, "x[%i,%i]" % (e, p)))
                        print("Do no reallocate: ", (e,p))
                    else:
                        y.append(model.NewIntVar(0, 0, "x[%i,%i]" % (e, p)))
            else:
                for e in all_employees:
                    y.append(model.NewIntVar(0, 1, "x[%i,%i]" % (e, p)))
        else:
            for e in all_employees:
                y.append(model.NewIntVar(0, 1, "x[%i,%i]" % (e, p)))
        print(y)
        x.append(y)
    # print(x)
    print("Initialising Constraints")

    # Every payroll assigned to at least one employee one day
    [model.Add(sum(x[p][e] for e in all_employees) == 1)
     for p in all_payrolls]

    # Constraint - sum of work across all working days <= max monthly hours for each employee
    [model.Add(sum(payrolls[p].get_processing_time() * x[p][e] for p in
                   all_payrolls) <= employees[e].get_max_hours()) for e in all_employees]

    # Constraint - payroll technicality <= employee technicality
    [model.Add(payrolls[p].get_technicality() * x[p][e] <= employees[e].get_technicality()) for p in
     all_payrolls for e in all_employees]

    print("Constraints intialised")

    model.Minimize(sum((employees[e].get_technicality()-payrolls[p].get_technicality()) * x[p][e] for e in all_employees for p in all_payrolls))
    # model.Maximize(sum((employees[e].get_max_hours()-sum(payrolls[p].get_processing_time() * x[p][e] for p in all_payrolls)) for e in all_employees))

    solver = cp_model.CpSolver()
    # solution_printer = PartialSolutionPrinter(x, num_employees, num_payrolls, range(10))
    # status = solver.SearchForAllSolutions(model, solution_printer)
    status = solver.Solve(model)

    print(status)

    print(time.time() - start_time)

    if status == cp_model.MODEL_SAT or status == cp_model.OPTIMAL:
        output = []
        count = 0
        for e in all_employees:
            employee_time_worked = 0
            print("Employee: ", employees[e].get_name(), " Technicality: ", employees[e].get_technicality())
            print("Assigned to: ")
            for p in all_payrolls:
                if solver.Value(x[p][e]) == 1:
                    count += 1
                    employee_time_worked += payrolls[p].get_processing_time()
                    print(payrolls[p].get_id(), ' due date', payrolls[p].get_due_date())
                    output.append([employees[e].get_name(), employees[e].get_technicality(), payrolls[p].get_id(), payrolls[p].get_technicality(), payrolls[p].get_processing_time(), payrolls[p].get_due_date()])
            print("Total work time: ", employee_time_worked, "Available work time: ", employees[e].get_max_hours())
        print(count)
        # print(output)
        output_df = pd.DataFrame(output,
                                 columns=['Employee', 'Employee Technicality', 'Payroll', 'Payroll Technicality',
                                          'Processing Time', 'Due date'])
        print(output_df)
        output_spreadsheet = pd.ExcelFile('Optimisation_Allocation.xlsx')
        output_df.to_excel(output_spreadsheet, "Allocation")
        print(solver.ObjectiveValue())

main()
