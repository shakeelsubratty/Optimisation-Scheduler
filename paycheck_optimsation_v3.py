# from __future__ import print_function
from ortools.sat.python import cp_model
import pandas as pd
import time
import numpy as np
import sys

from pandas import ExcelWriter
from openpyxl import load_workbook
from employee import Employee
from payroll import Payroll


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
    for p in range(start, end):
        due_date = input_payrolls.loc[p, 'Due date'].day
        due_dates.append(due_date)

    due_dates.sort()

    for date in due_dates:
        payrolls[date] = []

    for p in range(start, end):
        due_date = input_payrolls.loc[p, 'Due date'].day
        try:
            payroll = Payroll(input_payrolls.loc[p, 'Payroll'],
                              input_payrolls.loc[p, 'Prev. Employee'],
                              int(input_payrolls.loc[p, 'Technicality']), due_date,
                              input_payrolls.loc[p, 'Paydate'], input_payrolls.loc[p, 'Data sent date'],
                              input_payrolls.loc[p, 'Time in Minutes'],
                              input_payrolls.loc[p, 'Do not reallocate Flag'] == 'Y')
        except ValueError:
            print("VALUEERROR: ", input_payrolls.loc[p, 'Payroll'])
            return

        payrolls[due_date].append(payroll)
    return payrolls, len(input_payrolls)


def allocate(employees, payrolls, due_date, count):
    toggle_do_not_reallocate = False

    model = cp_model.CpModel()

    num_employees = len(employees)
    num_payrolls = len(payrolls)
    print(num_payrolls)
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

    print("Clearing Employee 2 day turnover period")
    for e in all_employees:
        employees[e].clear_allocated_payrolls_2_days_prior(due_date)

    print("Initialising Constraints")

    # Every payroll assigned to at least one employee
    [model.Add(sum(x[p][e] for e in all_employees) == 1)
     for p in all_payrolls]

    # employee_allocated_payroll_total_times = {}
    # for e in all_employees:
    #     employee_allocated_payroll_total_times[employees[e].get_name()] = employees[e].get_allocated_payrolls_total_time()

    employee_allocated_payroll_times = {}
    employee_possible_payroll_processing_times = {}
    employee_allocated_payroll_times_2days_from_due_date = {}
    for e in all_employees:
        employee_allocated_payroll_times[employees[e].get_name()] = employees[e].get_allocated_payrolls_total_time()
        employee_possible_payroll_processing_times[employees[e].get_name()] = sum(
            payrolls[p].get_processing_time() * x[p][e] for p in
            all_payrolls)
        employee_allocated_payroll_times_2days_from_due_date[employees[e].get_name()] = employees[
            e].get_allocated_payrolls_time_2days_from_due_date()

    # Constraint
    [model.Add((employee_possible_payroll_processing_times[employees[e].get_name()] + employee_allocated_payroll_times[
        employees[e].get_name()] <= employees[e].get_max_hours())) for e in all_employees]

    if (due_date == 21):
        print("p")
        [model.Add((employee_possible_payroll_processing_times[employees[e].get_name()] <= 1160 -
                    employee_allocated_payroll_times_2days_from_due_date[employees[e].get_name()])) for e in
         all_employees]
    else:
        [model.Add((employee_possible_payroll_processing_times[employees[e].get_name()] <= 840 -
                    employee_allocated_payroll_times_2days_from_due_date[employees[e].get_name()])) for e in
         all_employees]

    # Constraint - payroll technicality <= employee technicality
    [model.Add(payrolls[p].get_technicality() * x[p][e] <= employees[e].get_technicality()) for p in
     all_payrolls for e in all_employees]

    print("Constraints intialised")

    # model.Minimize(sum(
    #     (employees[e].get_technicality() - payrolls[p].get_technicality()) * x[p][e] for e in all_employees for p in
    #     all_payrolls))

    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    print(status)
    # count = 0
    if status == cp_model.MODEL_SAT or status == cp_model.OPTIMAL:
        output = []
        for e in all_employees:
            print("Employee: ", employees[e].get_name(), " Technicality: ", employees[e].get_technicality())
            print("Assigned to: ")
            allocated_payrolls = []
            for p in all_payrolls:
                if solver.Value(x[p][e]) == 1:
                    count += 1
                    allocated_payrolls.append(payrolls[p])
                    print(payrolls[p].get_id(), ' due date', payrolls[p].get_due_date())
                    output.append([employees[e].get_name(), employees[e].get_technicality(), payrolls[p].get_id(),
                                   payrolls[p].get_technicality(), payrolls[p].get_processing_time(),
                                   payrolls[p].get_due_date()])
            employees[e].allocate_payrolls(allocated_payrolls, due_date)
            print(employees[e].get_allocated_payrolls_total_time())
    return count


def main():
    input_spreadsheet = pd.ExcelFile('Optimisation_Spreadsheet_v3.xlsx')

    employees = import_employees(input_spreadsheet)
    start_time = time.time()
    count = 0
    due_dates = []

    payrolls, max_index = import_payrolls(input_spreadsheet, 0, 500)

    for date in payrolls:
        count = allocate(employees, payrolls[date], date, count)
        if date not in due_dates:
            due_dates.append(date)

    payrolls, max_index = import_payrolls(input_spreadsheet, 501, 1000)

    for date in payrolls:
        count = allocate(employees, payrolls[date], date, count)
        if date not in due_dates:
            due_dates.append(date)

    payrolls, max_index = import_payrolls(input_spreadsheet, 1001, 1234)

    for date in payrolls:
        count = allocate(employees, payrolls[date], date, count)
        if date not in due_dates:
            due_dates.append(date)

    print(time.time() - start_time)
    print("COUNT", count)

    # Allocation output
    output = []
    count = 0
    for e in employees:
        for p in e.get_allocated_payrolls():
            count += 1
            output.append([e.get_name(), e.get_technicality(), p.get_id(),
                           p.get_technicality(), p.get_processing_time(),
                           p.get_due_date()])
    print(count)
    output_df = pd.DataFrame(output,
                             columns=['Employee', 'Employee Technicality', 'Payroll', 'Payroll Technicality',
                                      'Processing Time', 'Due date'])
    print(output_df)

    # Employee/time surface output
    due_dates.sort()
    output_surface = []
    for date in due_dates:
        sum_of_mins = 0
        capacity_util = 0
        for e in employees:
            for p in e.get_allocated_payrolls():
                if p.get_due_date() == date:
                    sum_of_mins += p.get_processing_time()
        output_surface.append([date, sum_of_mins])
    output_surface_df = pd.DataFrame(output_surface, columns=['Due Date', 'Sum of mins'])

    with ExcelWriter('Optimisation_Allocation_v3.xlsx', engine='openpyxl') as writer:
        book = load_workbook('Optimisation_Allocation_v3.xlsx')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        output_df.to_excel(writer, "Allocation")
        output_surface_df.to_excel(writer, "Emp vs Time Surface")
        writer.save()

main()
