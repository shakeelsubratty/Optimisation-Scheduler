# from __future__ import print_function
from ortools.sat.python import cp_model
import numpy as np
#  - Possibly necessary imports for building with Windows
# import numpy.random.common
# import numpy.random.bounded_integers
# import numpy.random.entropy
import pandas as pd
import time
import sys
import easygui

from pandas import ExcelWriter
from openpyxl import load_workbook
from employee import Employee
from payroll import Payroll

# Paycheck optimisation script - uses a constraint programming solver to allocate payroll tasks to employees.

"""
Import Employees function - takes input_employees dataframe, creates and returns list of Employee objects.

Params: input_employees - input dataframe
Returns: employees - list of Employee objects
"""


def import_employees(input_employees):
    employees = []
    for e in range(len(input_employees)):
        try:
            employee = Employee(input_employees.loc[e, 'Employee'], input_employees.loc[e, 'Team'],
                                input_employees.loc[e, 'Role'], int(input_employees.loc[e, 'Technicality']),
                                int(input_employees.loc[e, 'Monthly Minutes']),
                                int(input_employees.loc[e, 'Availability in Minutes']))
        except ValueError:
            print("VALUEERROR: ", input_employees.loc[e, 'Employee'])
            easygui.msgbox(
                "An error has occured! Missing value for Employee: %s. "
                "\nScript terminating early, please fix and try again." % (
                    input_employees.loc[e, 'Employee']), "Missing value error")
            raise
        except KeyError:
            print("KEYERROR")
            easygui.msgbox(
                "An error has occured! Columns in Employee sheet incorrectly named. "
                "\nScript terminating early, please fix and try again.",
                "Incorrect column name error")
            raise
        employees.append(employee)
    return employees


"""
Import Payrolls function - takes input_payrolls dataframe between between start and end indices, and returns list of Payroll objects

Params:
	input_payrolls - input dataframe
	start - index (row) to start from in Payroll dataframe
	end - index (row) to end at in payroll dataframe
Returns:
	payrolls - list of Payrolls objects
	len(input_payrolls) - number of payrolls imported
	due_date_capacities - dictionary containing the assigned capacity per employee for each payroll due date
"""


def import_payrolls(input_payrolls, start, end):
    payrolls = {}  # Dictionary storing lists of Payrolls due on each due date
    due_dates = []
    due_date_capacities = {}

    # Find unique due dates between start and end date, sort dates ASC
    for p in range(start, end):
        try:
            due_date = input_payrolls.loc[p, 'Due date'].day
        except KeyError:
            print("KEYERROR")
            easygui.msgbox(
                "An error has occured! Due date column incorrectly named. "
                "\nScript terminating early, please fix and try again.",
                "Incorrect column name error")
            raise
        due_dates.append(due_date)
    due_dates.sort()

    # Use due dates as keys in payrolls dict, assigning empty array to each
    for date in due_dates:
        payrolls[date] = []

    # Pull unique due dates and capacities from payroll spreadsheet. Drop null values and convert to lists.
    try:
        unique_due_dates = input_payrolls['Unique Due Dates'].dropna().tolist()
        capacities = input_payrolls['Capacity per person'].dropna().tolist()
    except KeyError:
        print("KEYERROR")
        easygui.msgbox(
            "An error has occured! Unique Due Dates / Capacity per person columns incorrectly named. "
            "\nScript terminating early, please fix and try again.",
            "Incorrect column name error")
        raise

    # Check each due date has a capacity assigned
    if len(unique_due_dates) != len(capacities):
        print("TABLEERROR")
        easygui.msgbox(
            "An error has occured! Missing values for Unique Due Dates / Capacities. "
            "\nScript terminating early, please fix and try again.",
            "Missing table entry error")
        return

    # Assign each capacity to appropriate due date in due_date_capacities dictionary
    for i in range(0, len(capacities)):
        due_date_capacities[unique_due_dates[i].day] = int(capacities[i])

    # Check each due date that exists in main payroll table has been given a capacity
    if all(elem in payrolls.keys() for elem in due_date_capacities.keys()):
        print("TABLEERROR")
        easygui.msgbox(
            "An error has occured! Missing values for Unique Due Dates / Capacities. "
            "\nScript terminating early, please fix and try again.",
            "Missing table entry error")
        return

    print(due_date_capacities)

    # For each payroll entry in range, attempt to create Payroll object and append to list on payrolls dictionary at that payroll's due date
    for p in range(start, end):
        try:
            due_date = int(input_payrolls.loc[p, 'Due date'].day)
            payroll = Payroll(input_payrolls.loc[p, 'Payroll'],
                              input_payrolls.loc[p, 'Prev. Employee'],
                              int(input_payrolls.loc[p, 'Technicality']), due_date,
                              input_payrolls.loc[p, 'Paydate'], input_payrolls.loc[p, 'Data sent date'],
                              int(input_payrolls.loc[p, 'Time in Minutes']),
                              input_payrolls.loc[p, 'Do not reallocate Flag'] == 'Y')
        except ValueError:
            print("VALUEERROR: ", input_payrolls.loc[p, 'Payroll'])
            easygui.msgbox(
                "An error has occured! Missing value for Payroll: %s. "
                "\nScript terminating early, please fix and try again." % (
                    input_payrolls.loc[p, 'Payroll']), "Missing value error")
            raise
        except KeyError:
            print("KEYERROR")
            easygui.msgbox(
                "An error has occured! Columns in Payroll sheet incorrectly named. "
                "\nScript terminating early, please fix and try again.",
                "Incorrect column name error")
            raise

        payrolls[due_date].append(payroll)
    return payrolls, len(input_payrolls), due_date_capacities


"""
Allocate function - constructs Constraint Progromming model, establishes constraints and performs allocation procedure.
	Allocation is carried out for each individual due date seperately, given the lists of employees and payrolls due on that day and
	a two day turnover period prior to the current due date.
	Allocation is stored in Employee local fields.

Params:
	employees - list of Employees
	payrolls - list of Payrolls due on that date
	due_date_capacities - dictionary containing the assigned capacity per employee for each payroll due date
	due_date - the due date of the payrolls being allocated
	count - maintains a integer count of payrolls processed
Returns:
	count - maintains a integer count of payrolls processed
"""


def allocate(employees, payrolls, due_date_capacities, due_date, count):
    toggle_do_not_reallocate = False  # Toggle whether to accept values from "Do not reallocate" column in Payrolls spreadsheet.

    model = cp_model.CpModel()  # Constraint programming model and solver

    # Construct 2-D search space for allocations
    num_employees = len(employees)
    num_payrolls = len(payrolls)
    all_employees = range(num_employees)
    all_payrolls = range(num_payrolls)
    x = []  # For each payroll, stores a possible employee
    for p in all_payrolls:
        y = []  # For each employee, stores CP model integer variable between (min, max). Solver adjusts these values between min/max to find a solution that satisfies constraints.
        if toggle_do_not_reallocate:
            if payrolls[p].get_do_not_reallocate():
                prev_emp_str = payrolls[p].get_previous_employee()
                for e in all_employees:
                    if employees[e].get_name() == prev_emp_str:
                        y.append(model.NewIntVar(1, 1, "x[%i,%i]" % (e,
                                                                     p)))  # If get_do_not_reallocate() == True, fix possible values of assignment for current employee to that payroll to (1,1) i.e. value must be 1.
                        print("Do no reallocate: ", (e, p))
                    else:
                        y.append(model.NewIntVar(0, 0, "x[%i,%i]" % (e, p)))  # For all other employees, fix value to 0
            else:
                for e in all_employees:  # If get_do_not_reallocate() == False, allow possible values of assignment to be 0 or 1.
                    y.append(model.NewIntVar(0, 1, "x[%i,%i]" % (e, p)))
        else:
            for e in all_employees:
                y.append(model.NewIntVar(0, 1, "x[%i,%i]" % (e,
                                                             p)))  # If toggle_do_not_reallocate == False,  allow possible values of assignment to be 0 or 1 for all payrolls.
        print(y)
        x.append(y)

    # Initialising Constraints
    print("Initialising Constraints")

    # Remove any payrolls that are now outside each employee's two day turnover period
    for e in all_employees:
        employees[e].clear_allocated_payrolls_2_days_prior(due_date)
    print("Cleared employee two day turnover period")


    # Constraint - Every payroll assigned to at least one employee
    [model.Add(sum(x[p][e] for e in all_employees) == 1)
     for p in all_payrolls]

    # Preparing repeatedly accessed dictionaries for easy access
    employee_allocated_payroll_times = {}  # Total time for currently allocated payrolls for each employee
    employee_possible_payroll_processing_times = {}  # Total time for payrolls that could be assigned for each employee
    employee_allocated_payroll_times_2days_from_due_date = {}  # Total time for currently allocated payrolls within two day period for each employee
    for e in all_employees:
        employee_allocated_payroll_times[employees[e].get_name()] = employees[e].get_allocated_payrolls_total_time()
        employee_possible_payroll_processing_times[employees[e].get_name()] = sum(
        payrolls[p].get_processing_time() * x[p][e] for p in
            all_payrolls)
        employee_allocated_payroll_times_2days_from_due_date[employees[e].get_name()] = employees[
            e].get_allocated_payrolls_time_2days_from_due_date()


    # Constraint - For each employee, possible processing times + currently allocated payroll times must be <= the employee's maxmimum amount of working time
    [model.Add((employee_possible_payroll_processing_times[employees[e].get_name()] + employee_allocated_payroll_times[
        employees[e].get_name()] <= employees[e].get_max_hours())) for e in all_employees]

    # Constraint - For each employee, possible processing times must be <= capacity for current due date - processing time
    #	for payrolls assigned to that employee over two day turnover period.
    [model.Add((employee_possible_payroll_processing_times[employees[e].get_name()] <= due_date_capacities[due_date] -
                employee_allocated_payroll_times_2days_from_due_date[employees[e].get_name()])) for e in
     all_employees]

    # Constraint - For each employee, payroll technicality <= employee technicality
    [model.Add(payrolls[p].get_technicality() * x[p][e] <= employees[e].get_technicality()) for p in
     all_payrolls for e in all_employees]

    print("Constraints intialised")

    # Find a solution to the model
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    # If solution found, append assignments to Employee object fields by calling <Employee>.allocate_payrolls(allocate_payrolls, due_date)
    print(status)
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

"""
Main function - reads input spreadsheets, carries out allocation procedure in blocks of 500 payrolls, and writes to output spreadsheets.

"""


def main():
    # Read input spreadsheet
    try:
        input_spreadsheet = pd.ExcelFile('Optimisation_Spreadsheet.xlsx')
    except FileNotFoundError:
        print("FileNotFound")
        easygui.msgbox(
            "An error has occured! Input spreadsheet has not been found. "
            "\nEnsure a file named %s exists in the same directory as the script and try again." % (
                "Optimisation_Spreadsheet_v3.xlsx"), "File not found error")
        raise
    # Open output spreadsheet for writing
    try:
        writer = ExcelWriter('Optimisation_Allocation.xlsx', engine='openpyxl')
        book = load_workbook('Optimisation_Allocation.xlsx')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        # Clear any existing values in columns
        df_empty_allocation = pd.DataFrame(
            columns=['Employee', 'Employee Technicality', 'Payroll', 'Payroll Technicality',
                     'Processing Time', 'Due date'])
        df_empty_allocation.to_excel(writer, "Allocation")
        df_empty_surface = pd.DataFrame(columns=['Due Date', 'Sum of mins', 'Capacity', 'Capacity Utilisation'])
        df_empty_surface.to_excel(writer, "Emp vs Time Surface")

    except (FileNotFoundError, KeyError):
        print("FileNotFound")
        easygui.msgbox(
            "An error has occured! Output spreadsheet has not been found. "
            "\nEnsure a file named %s exists in the same directory as the script and try again." % (
                "Optimisation_Allocation_v3.xlsx"), "File not found error")
        raise

    input_employees = pd.read_excel(input_spreadsheet, 'Employees')  # Reads from "Employee" worksheet
    input_payrolls = pd.read_excel(input_spreadsheet, 'Payrolls')  # Reads from "Payrolls" worksheet

    employees = import_employees(input_employees)
    start_time = time.time()
    count = 0
    due_dates = []
    total_number_of_payrolls = len(input_payrolls)


    # Perform Allocation procedure in blocks of 500 payrolls at a time
    for index in range(0, total_number_of_payrolls):
        if index == 0:
            payrolls, max_index, due_date_capacities = import_payrolls(input_payrolls, index, index + 500)
            for date in payrolls:
                count = allocate(employees, payrolls[date], due_date_capacities, date, count)
                if date not in due_dates:
                    due_dates.append(date)
        if index % 500 == 1 and index > 1:
            end = index + 499 if ((index + 499) < total_number_of_payrolls) else total_number_of_payrolls
            payrolls, max_index, due_date_capacities = import_payrolls(input_payrolls, index, end)
            for date in payrolls:
                count = allocate(employees, payrolls[date], due_date_capacities, date, count)
                if date not in due_dates:
                    due_dates.append(date)

    print("Time taken: ", time.time() - start_time)
    print("COUNT: ", count, "Number of payrolls: ", total_number_of_payrolls)

    # Allocation output
    output = []
    output_payroll_ids = []
    for e in employees:
        for p in e.get_allocated_payrolls():
            output.append([e.get_name(), e.get_technicality(), p.get_id(),
                           p.get_technicality(), p.get_processing_time(),
                           p.get_due_date()])
            output_payroll_ids.append(p.get_id())

    # Check to see if any payrolls were not allocated
    failed_to_allocate = []
    for elem in input_payrolls['Payroll']:
        if elem not in output_payroll_ids:
            failed_to_allocate.append(elem)

    if len(failed_to_allocate) > 0:
        msg = "The following payrolls could not be allocated: " + str(failed_to_allocate) \
              + " \nPlease try again or allocate them manually."
        easygui.msgbox(msg)
    output_df = pd.DataFrame(output,
                             columns=['Employee', 'Employee Technicality', 'Payroll', 'Payroll Technicality',
                                      'Processing Time', 'Due date'])
    print(output_df)

    # Employee/time surface output
    due_dates.sort()
    output_surface = []
    sum_of_mins_history = {}

    for date in due_dates:
        sum_of_mins = 0
        for e in employees:
            for p in e.get_allocated_payrolls():
                if p.get_due_date() == date:
                    sum_of_mins += p.get_processing_time()
        sum_of_mins_history[date] = sum_of_mins
        if date - 1 in sum_of_mins_history.keys():
            prior = sum_of_mins_history[date - 1]
        else:
            prior = 0

        capacity = due_date_capacities[date] * len(employees) - prior
        capacity_util = int(sum_of_mins / capacity * 100)
        output_surface.append([date, sum_of_mins, capacity, capacity_util])
    output_surface_df = pd.DataFrame(output_surface,
                                     columns=['Due Date', 'Sum of mins', 'Capacity', 'Capacity Utilisation'])

    output_df.to_excel(writer, "Allocation")
    output_surface_df.to_excel(writer, "Emp vs Time Surface")
    writer.save()

main()
input('..')
