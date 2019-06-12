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

def import_employees(input_spreadsheet):
    input_employees = pd.read_excel(input_spreadsheet, 'Employees')
    employees = []
    for e in range(len(input_employees)):
        employee = Employee(input_employees.loc[e, 'Employee'], input_employees.loc[e, 'Team'],
                            input_employees.loc[e, 'Role'], input_employees.loc[e, 'Technicality'],
                            input_employees.loc[e, 'Monthly Minutes'],
                            input_employees.loc[e, 'Availability in Minutes'])
        employees.append(employee)
    return employees

def import_payrolls(input_spreadsheet, start, end):
    input_payrolls = pd.read_excel(input_spreadsheet, 'Payrolls')
    payrolls = {}
    due_dates = []
    for p in range(start,end):
        due_date = input_payrolls.loc[p, 'Due date'].day
        due_dates.append(due_date)

    due_dates.sort()

    for date in due_dates:
        payrolls[date] = []

    for p in range(start,end):
        due_date = input_payrolls.loc[p, 'Due date'].day
        payroll = Payroll(input_payrolls.loc[p, 'Payroll'],
                          input_payrolls.loc[p, 'Prev. Employee'],
                          input_payrolls.loc[p, 'Technicality'], due_date,
                          input_payrolls.loc[p, 'Paydate'], input_payrolls.loc[p, 'Data sent date'],
                          input_payrolls.loc[p, 'Time in Minutes'],
                          input_payrolls.loc[p, 'Do not reallocate Flag'] == 'Y')
        payrolls[due_date].append(payroll)
    return payrolls, len(input_payrolls)

def allocate(employees, payrolls, due_date):

    toggle_do_not_reallocate = False

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
                        print("Do no reallocate: ", (e, p))
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

    # employee_allocated_payroll_total_times = {}
    # for e in all_employees:
    #     employee_allocated_payroll_total_times[employees[e].get_name()] = employees[e].get_allocated_payrolls_total_time()

    employee_allocated_payroll_times = {}
    employee_possible_payroll_processing_times = {}
    for e in all_employees:
        employee_allocated_payroll_times[employees[e].get_name()] = employees[e].get_allocated_payrolls_total_time()
        employee_possible_payroll_processing_times[employees[e].get_name()] = sum(payrolls[p].get_processing_time() * x[p][e] for p in
                   all_payrolls)

    # Constraint
    [model.Add((employee_possible_payroll_processing_times[employees[e].get_name()] + employee_allocated_payroll_times[employees[e].get_name()] <= employees[e].get_max_hours())) for e in all_employees]

    [model.Add((employee_possible_payroll_processing_times[employees[e].get_name()] <= 2100 - employees[e].get_allocated_payrolls_time_7days_from_due_date())) for e in all_employees]

    # Constraint -
    # [model.Add((employee_possible_payroll_processing_times[employees[e].get_name()] + employee_allocated_payroll_times[employees[e].get_name()] <= employees[e].get_max_hours()) and (employee_possible_payroll_processing_times[employees[e].get_name()] <= 2100 - employees[e].get_allocated_payrolls_time_7days_from_due_date(due_date))) for e in all_employees]

    # Constraint - payroll technicality <= employee technicality
    [model.Add(payrolls[p].get_technicality() * x[p][e] <= employees[e].get_technicality()) for p in
     all_payrolls for e in all_employees]

    print("Constraints intialised")

    # model.Minimize(sum(
    #     (employees[e].get_technicality() - payrolls[p].get_technicality()) * x[p][e] for e in all_employees for p in
    #     all_payrolls))

    solver = cp_model.CpSolver()
    # solution_printer = PartialSolutionPrinter(x, num_employees, num_payrolls, range(10))
    # status = solver.SearchForAllSolutions(model, solution_printer)
    status = solver.Solve(model)

    print(status)

    if status == cp_model.MODEL_SAT or status == cp_model.OPTIMAL:
        output = []
        for e in all_employees:
            print("Employee: ", employees[e].get_name(), " Technicality: ", employees[e].get_technicality())
            print("Assigned to: ")
            allocated_payrolls = []
            for p in all_payrolls:
                if solver.Value(x[p][e]) == 1:
                    allocated_payrolls.append(payrolls[p])
                    print(payrolls[p].get_id(), ' due date', payrolls[p].get_due_date())
                    output.append([employees[e].get_name(), employees[e].get_technicality(), payrolls[p].get_id(),
                                   payrolls[p].get_technicality(), payrolls[p].get_processing_time(),
                                   payrolls[p].get_due_date()])
            employees[e].allocate_payrolls(allocated_payrolls, due_date)
            print(employees[e].get_allocated_payrolls_total_time())


def main():
    input_spreadsheet = pd.ExcelFile('Optimisation_Spreadsheet_v1.xlsx')

    employees = import_employees(input_spreadsheet)
    start_time = time.time()


    payrolls, max_index = import_payrolls(input_spreadsheet, 1, 500)

    for date in payrolls:
        allocate(employees, payrolls[date], date)

    payrolls, max_index = import_payrolls(input_spreadsheet, 501, 1000)

    for date in payrolls:
        allocate(employees, payrolls[date], date)

    payrolls, max_index = import_payrolls(input_spreadsheet, 1001, 1319)

    for date in payrolls:
        allocate(employees, payrolls[date], date)

    print(time.time() - start_time)

    # allocate(employees, payrolls[1])
    # allocate(employees, payrolls[5])

main()