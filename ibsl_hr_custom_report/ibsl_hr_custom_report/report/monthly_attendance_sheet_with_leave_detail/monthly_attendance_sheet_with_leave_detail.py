# Copyright (c) 2015, Frappe Technologies Pvt. Ltd. and Contributors
# License: GNU General Public License v3. See license.txt


from calendar import monthrange
from itertools import groupby
from typing import Dict, List, Optional, Tuple
from datetime import datetime
import frappe
from frappe import _
from frappe.query_builder.functions import Count, Extract, Sum
from frappe.utils import cint, cstr, getdate

Filters = frappe._dict

status_map = {
    "Present": "P",
    "Absent": "A",
    "Half Day": "HD",
    "Work From Home": "WFH",
    "On Leave": "L",
    "Holiday": "H",
    "Weekly Off": "WO",
}

day_abbr = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

def extract_time(datetime_value) -> str:
    """Extracts and returns the time part from a datetime object or string."""
    if isinstance(datetime_value, datetime):
        return datetime_value.strftime("%H:%M:%S")
    elif isinstance(datetime_value, str) and datetime_value.strip():  # Check if string is not empty
        return datetime.strptime(datetime_value, "%Y-%m-%d %H:%M:%S").strftime("%H:%M:%S")
    return ""
def execute(filters: Optional[Filters] = None) -> Tuple:
	filters = frappe._dict(filters or {})

	if not (filters.month and filters.year):
		frappe.throw(_("Please select month and year."))

	attendance_map = get_attendance_map(filters)
	if not attendance_map:
		frappe.msgprint(_("No attendance records found."), alert=True, indicator="orange")
		return [], [], None, None

	columns = get_columns(filters)
	data = get_data(filters, attendance_map)

	if not data:
		frappe.msgprint(
			_("No attendance records found for this criteria."), alert=True, indicator="orange"
		)
		return columns, [], None, None

	message = get_message() if not filters.summarized_view else ""
	#chart = get_chart_data(attendance_map, filters)

	return columns, data, message


def get_message() -> str:
	message = ""
	colors = ["green", "red", "orange", "green", "#318AD8", "", ""]

	count = 0
	for status, abbr in status_map.items():
		message += f"""
			<span style='border-left: 2px solid {colors[count]}; padding-right: 12px; padding-left: 5px; margin-right: 3px;'>
				{status} - {abbr}
			</span>
		"""
		count += 1

	return message


def get_columns(filters: Filters) -> List[Dict]:
	columns = []

	if filters.group_by:
		options_mapping = {
			"Branch": "Branch",
			"Grade": "Employee Grade",
			"Department": "Department",
			"Designation": "Designation",
		}
		options = options_mapping.get(filters.group_by)
		columns.append(
			{
				"label": _(filters.group_by),
				"fieldname": frappe.scrub(filters.group_by),
				"fieldtype": "Link",
				"options": options,
				"width": 120,
			}
		)

	columns.extend(
		[
			{
				"label": _("Employee"),
				"fieldname": "employee",
				"fieldtype": "Link",
				"options": "Employee",
				"width": 135,
			},
			{"label": _("Employee Name"), "fieldname": "employee_name", "fieldtype": "Data", "width": 120},
		]
	)

	if filters.summarized_view:
		columns.extend(
			[
				{
					"label": _("Total Present"),
					"fieldname": "total_present",
					"fieldtype": "Float",
					"width": 110,
				},
				{"label": _("Total Leaves"), "fieldname": "total_leaves", "fieldtype": "Float", "width": 110},
				{"label": _("Total Absent"), "fieldname": "total_absent", "fieldtype": "Float", "width": 110},
				{
					"label": _("Total Holidays"),
					"fieldname": "total_holidays",
					"fieldtype": "Float",
					"width": 120,
				},
				{
					"label": _("Unmarked Days"),
					"fieldname": "unmarked_days",
					"fieldtype": "Float",
					"width": 130,
				},
			]
		)
		columns.extend(get_columns_for_leave_types())
		columns.extend(
			[
				{
					"label": _("Total Late Entries"),
					"fieldname": "total_late_entries",
					"fieldtype": "Float",
					"width": 140,
				},
				{
					"label": _("Total Early Exits"),
					"fieldname": "total_early_exits",
					"fieldtype": "Float",
					"width": 140,
				},
			]
		)
	else:
		columns.append({"label": _("Shift"), "fieldname": "shift", "fieldtype": "Data", "width": 120})
		columns.extend(get_columns_for_days(filters))

	return columns


def get_columns_for_leave_types() -> List[Dict]:
	leave_types = frappe.db.get_all("Leave Type", pluck="name")
	types = []
	for entry in leave_types:
		types.append(
			{"label": entry, "fieldname": frappe.scrub(entry), "fieldtype": "Float", "width": 120}
		)

	return types


def get_columns_for_days(filters: Filters) -> List[Dict]:
    total_days = get_total_days_in_month(filters)
    days = []

    for day in range(1, total_days + 1):
        day_str = cstr(day)
        date = "{}-{}-{}".format(cstr(filters.year), cstr(filters.month), day)
        weekday = day_abbr[getdate(date).weekday()]
        label = "{} {}".format(day, weekday)
        days.append({"label": label, "fieldtype": "Data", "fieldname": day_str, "width": 125})
        #days.append({"label": f"{label} In Time", "fieldtype": "Data", "fieldname": f"in_time_{day_str}", "width": 80})
        #days.append({"label": f"{label} Out Time", "fieldtype": "Data", "fieldname": f"out_time_{day_str}", "width": 80})
        #days.append({"label": f"{label} Leave Type", "fieldtype": "Data", "fieldname": f"leave_type{day_str}", "width": 150})

    return days



def get_total_days_in_month(filters: Filters) -> int:
	return monthrange(cint(filters.year), cint(filters.month))[1]


def get_data(filters: Filters, attendance_map: Dict) -> List[Dict]:
	employee_details, group_by_param_values = get_employee_related_details(filters)
	holiday_map = get_holiday_map(filters)
	data = []

	if filters.group_by:
		group_by_column = frappe.scrub(filters.group_by)

		for value in group_by_param_values:
			if not value:
				continue

			records = get_rows(employee_details[value], filters, holiday_map, attendance_map)

			if records:
				data.append({group_by_column: value})
				data.extend(records)
	else:
		data = get_rows(employee_details, filters, holiday_map, attendance_map)

	return data


def get_attendance_map(filters: Filters) -> Dict:
    attendance_list = get_attendance_records(filters)
    attendance_map = {}
    leave_map = {}

    for d in attendance_list:
        if d.status == "On Leave":
            leave_map.setdefault(d.employee, []).append({
                "day": d.day_of_month,
                "leave_type": d.leave_type
            })
            continue

        if d.shift is None:
            d.shift = ""

        attendance_map.setdefault(d.employee, {}).setdefault(d.shift, {})
        if d.status == "Present":
            attendance_map[d.employee][d.shift][d.day_of_month] = {
                "status": d.status,
                "in_time": d.in_time,
                "out_time": d.out_time
            }
        else:
            attendance_map[d.employee][d.shift][d.day_of_month] = {"status": d.status}

    # leave is applicable for the entire day, so all shifts should show the leave entry
    for employee, leave_entries in leave_map.items():
        # no attendance records exist except leaves
        if employee not in attendance_map:
            attendance_map.setdefault(employee, {}).setdefault(None, {})

        for entry in leave_entries:
            day = entry["day"]
            leave_type = entry["leave_type"]
            for shift in attendance_map[employee].keys():
                attendance_map[employee][shift][day] = {
                    "status": "On Leave",
                    "leave_type": leave_type
                }

    return attendance_map




def get_attendance_records(filters: Filters) -> List[Dict]:
    Attendance = frappe.qb.DocType("Attendance")
    query = (
        frappe.qb.from_(Attendance)
        .select(
            Attendance.employee,
            Extract("day", Attendance.attendance_date).as_("day_of_month"),
            Attendance.status,
            Attendance.in_time,
            Attendance.out_time,
            Attendance.shift,
            Attendance.leave_type
        )
        .where(
            (Attendance.docstatus == 1)
            & (Attendance.company == filters.company)
            & (Extract("month", Attendance.attendance_date) == filters.month)
            & (Extract("year", Attendance.attendance_date) == filters.year)
        )
    )

    if filters.employee:
        query = query.where(Attendance.employee == filters.employee)
    query = query.orderby(Attendance.employee, Attendance.attendance_date)

    return query.run(as_dict=1)



def get_employee_related_details(filters: Filters) -> Tuple[Dict, List]:
	"""Returns
	1. nested dict for employee details
	2. list of values for the group by filter
	"""
	Employee = frappe.qb.DocType("Employee")
	query = (
		frappe.qb.from_(Employee)
		.select(
			Employee.name,
			Employee.employee_name,
			Employee.designation,
			Employee.grade,
			Employee.department,
			Employee.branch,
			Employee.company,
			Employee.holiday_list,
		)
		.where(Employee.company == filters.company)
	)

	if filters.employee:
		query = query.where(Employee.name == filters.employee)

	group_by = filters.group_by
	if group_by:
		group_by = group_by.lower()
		query = query.orderby(group_by)

	employee_details = query.run(as_dict=True)

	group_by_param_values = []
	emp_map = {}

	if group_by:
		for parameter, employees in groupby(employee_details, key=lambda d: d[group_by]):
			group_by_param_values.append(parameter)
			emp_map.setdefault(parameter, frappe._dict())

			for emp in employees:
				emp_map[parameter][emp.name] = emp
	else:
		for emp in employee_details:
			emp_map[emp.name] = emp

	return emp_map, group_by_param_values


def get_holiday_map(filters: Filters) -> Dict[str, List[Dict]]:
	"""
	Returns a dict of holidays falling in the filter month and year
	with list name as key and list of holidays as values like
	{
	        'Holiday List 1': [
	                {'day_of_month': '0' , 'weekly_off': 1},
	                {'day_of_month': '1', 'weekly_off': 0}
	        ],
	        'Holiday List 2': [
	                {'day_of_month': '0' , 'weekly_off': 1},
	                {'day_of_month': '1', 'weekly_off': 0}
	        ]
	}
	"""
	# add default holiday list too
	holiday_lists = frappe.db.get_all("Holiday List", pluck="name")
	default_holiday_list = frappe.get_cached_value("Company", filters.company, "default_holiday_list")
	holiday_lists.append(default_holiday_list)

	holiday_map = frappe._dict()
	Holiday = frappe.qb.DocType("Holiday")

	for d in holiday_lists:
		if not d:
			continue

		holidays = (
			frappe.qb.from_(Holiday)
			.select(Extract("day", Holiday.holiday_date).as_("day_of_month"), Holiday.weekly_off)
			.where(
				(Holiday.parent == d)
				& (Extract("month", Holiday.holiday_date) == filters.month)
				& (Extract("year", Holiday.holiday_date) == filters.year)
			)
		).run(as_dict=True)

		holiday_map.setdefault(d, holidays)

	return holiday_map


def get_rows(
	employee_details: Dict, filters: Filters, holiday_map: Dict, attendance_map: Dict
) -> List[Dict]:
	records = []
	default_holiday_list = frappe.get_cached_value("Company", filters.company, "default_holiday_list")

	for employee, details in employee_details.items():
		emp_holiday_list = details.holiday_list or default_holiday_list
		holidays = holiday_map.get(emp_holiday_list)

		if filters.summarized_view:
			attendance = get_attendance_status_for_summarized_view(employee, filters, holidays)
			if not attendance:
				continue

			leave_summary = get_leave_summary(employee, filters)
			entry_exits_summary = get_entry_exits_summary(employee, filters)

			row = {"employee": employee, "employee_name": details.employee_name}
			set_defaults_for_summarized_view(filters, row)
			row.update(attendance)
			row.update(leave_summary)
			row.update(entry_exits_summary)

			records.append(row)
		else:
			employee_attendance = attendance_map.get(employee)
			if not employee_attendance:
				continue

			attendance_for_employee = get_attendance_status_for_detailed_view(
				employee, filters, employee_attendance, holidays
			)
			# set employee details in the first row
			attendance_for_employee[0].update(
				{"employee": employee, "employee_name": details.employee_name}
			)

			records.extend(attendance_for_employee)

	return records


def set_defaults_for_summarized_view(filters, row):
	for entry in get_columns(filters):
		if entry.get("fieldtype") == "Float":
			row[entry.get("fieldname")] = 0.0


def get_attendance_status_for_summarized_view(
	employee: str, filters: Filters, holidays: List
) -> Dict:
	"""Returns dict of attendance status for employee like
	{'total_present': 1.5, 'total_leaves': 0.5, 'total_absent': 13.5, 'total_holidays': 8, 'unmarked_days': 5}
	"""
	summary, attendance_days = get_attendance_summary_and_days(employee, filters)
	if not any(summary.values()):
		return {}

	total_days = get_total_days_in_month(filters)
	total_holidays = total_unmarked_days = 0

	for day in range(1, total_days + 1):
		if day in attendance_days:
			continue

		status = get_holiday_status(day, holidays)
		if status in ["Weekly Off", "Holiday"]:
			total_holidays += 1
		elif not status:
			total_unmarked_days += 1

	return {
		"total_present": summary.total_present + summary.total_half_days,
		"total_leaves": summary.total_leaves + summary.total_half_days,
		"total_absent": summary.total_absent,
		"total_holidays": total_holidays,
		"unmarked_days": total_unmarked_days,
	}


def get_attendance_summary_and_days(employee: str, filters: Filters) -> Tuple[Dict, List]:
	Attendance = frappe.qb.DocType("Attendance")

	present_case = (
		frappe.qb.terms.Case()
		.when(((Attendance.status == "Present") | (Attendance.status == "Work From Home")), 1)
		.else_(0)
	)
	sum_present = Sum(present_case).as_("total_present")

	absent_case = frappe.qb.terms.Case().when(Attendance.status == "Absent", 1).else_(0)
	sum_absent = Sum(absent_case).as_("total_absent")

	leave_case = frappe.qb.terms.Case().when(Attendance.status == "On Leave", 1).else_(0)
	sum_leave = Sum(leave_case).as_("total_leaves")

	half_day_case = frappe.qb.terms.Case().when(Attendance.status == "Half Day", 0.5).else_(0)
	sum_half_day = Sum(half_day_case).as_("total_half_days")

	summary = (
		frappe.qb.from_(Attendance)
		.select(
			sum_present,
			sum_absent,
			sum_leave,
			sum_half_day,
		)
		.where(
			(Attendance.docstatus == 1)
			& (Attendance.employee == employee)
			& (Attendance.company == filters.company)
			& (Extract("month", Attendance.attendance_date) == filters.month)
			& (Extract("year", Attendance.attendance_date) == filters.year)
		)
	).run(as_dict=True)

	days = (
		frappe.qb.from_(Attendance)
		.select(Extract("day", Attendance.attendance_date).as_("day_of_month"))
		.distinct()
		.where(
			(Attendance.docstatus == 1)
			& (Attendance.employee == employee)
			& (Attendance.company == filters.company)
			& (Extract("month", Attendance.attendance_date) == filters.month)
			& (Extract("year", Attendance.attendance_date) == filters.year)
		)
	).run(pluck=True)

	return summary[0], days

from datetime import datetime
import calendar
def calculate_working_hours(in_time, out_time):
    """Returns working hours as HH:MM from in_time and out_time."""
    if not in_time or not out_time:
        return ""
    try:
        in_dt = in_time if isinstance(in_time, datetime) else datetime.strptime(in_time, "%Y-%m-%d %H:%M:%S")
        out_dt = out_time if isinstance(out_time, datetime) else datetime.strptime(out_time, "%Y-%m-%d %H:%M:%S")
        delta = out_dt - in_dt
        total_seconds = delta.total_seconds()
        hours = int(total_seconds // 3600)
        minutes = int((total_seconds % 3600) // 60)
        return f"{hours:02}:{minutes:02}"
    except Exception:
        return ""

def get_attendance_status_for_detailed_view(
    employee: str, filters: Filters, employee_attendance: Dict, holidays: List
) -> List[Dict]:
    """Returns list of shift-wise attendance status for employee
    [
            {'shift': 'Morning Shift', 1: 'A', 2: 'P', 3: 'A'...., 'in_time_1': '09:00', 'out_time_1': '18:00', ...},
            {'shift': 'Evening Shift', 1: 'P', 2: 'A', 3: 'P'...., 'in_time_2': '09:00', 'out_time_2': '18:00', ...}
    ]
    """
    total_days = get_total_days_in_month(filters)
    year = int(filters.year)
    month_num = int(filters.month)
    attendance_values = []
    #frappe.throw(str(filters.month))

    for shift, status_dict in employee_attendance.items():
        row = {"shift": shift}

        for day in range(1, total_days + 1):
            day_str = cstr(day)
            status_info = status_dict.get(day)
            #frappe.throw(str(status_info))
            if status_info is None and holidays:
                status_info = {'status': get_holiday_status(day, holidays)}

            if isinstance(status_info, dict):
                status = status_info.get('status')
                #da=status_info.get("name")
                #frappe.throw(str(da))
                #wo=status_info.get('in_time')
                in_time = status_info.get('in_time')
                out_time = status_info.get('out_time')
                #frappe.throw(str(status_info))
                #in_time = extract_time(status_info.get('in_time', ''))
                #out_time = extract_time(status_info.get('out_time', ''))
                leave_type = status_info.get('leave_type', '')
                #att_doc = frappe.get_value("Attendance", {"employee": employee, "attendance_date": f"{year}-{month:02d}-{int(day_str):02d}"}, ["in_time", "out_time"], as_dict=True)
                abbr = status_map.get(status, "")
                row[day_str] = abbr
                if abbr == "HD" and (not in_time or not out_time):
                    att_doc = frappe.get_value(
                        "Attendance",
                        {
                            "employee": employee,
                            "attendance_date": f"{year}-{month_num:02d}-{int(day_str):02d}"
                        },
                        ["in_time", "out_time"],
                        as_dict=True
                    )
                    if att_doc:
                        in_time = att_doc.in_time
                        out_time = att_doc.out_time

                # Calculate working hours
                wo = calculate_working_hours(in_time, out_time)
                #wo = calculate_working_hours(in_time, out_time)
                #row[f"in_time_{day_str}"] = in_time
                if leave_type:
                    cell_value = leave_type
                else:
                    if abbr == "HD":
                        cell_value = f"{abbr} - {wo if wo else '04:00'} hrs"
                    else:
                        cell_value = f"{abbr} - {wo} hrs" if wo else abbr

                # Apply color coding
                if abbr == "P":
                    cell_value = f'<span style="color: green;">{cell_value}</span>'
                elif abbr == "A":
                    cell_value = f'<span style="color: red;">{cell_value}</span>'
                elif abbr == "L":
                    cell_value = f'<span style="color: #4682b4;">{cell_value}</span>'
                elif abbr == "HD":
                    cell_value = f'<span style="color: orange;">{cell_value}</span>'

                row[day_str] = cell_value
                #if leave_type:
                	#row[day_str] = leave_type
                #else:
                	#row[day_str] = f"{abbr} - {wo} hrs" if wo else abbr
                #row[f"out_time_{day_str}"] = out_time
            else:
                status = status_info
                abbr = status_map.get(status, "")
                row[day_str] = abbr
                #row[f"in_time_{day_str}"] = ""
                #row[f"leave_type{day_str}"] = ""
                #row[f"out_time_{day_str}"] = ""

        attendance_values.append(row)

    return attendance_values


def get_holiday_status(day: int, holidays: List) -> str:
	status = None
	if holidays:
		for holiday in holidays:
			if day == holiday.get("day_of_month"):
				if holiday.get("weekly_off"):
					status = "Weekly Off"
				else:
					status = "Holiday"
				break
	return status


def get_leave_summary(employee: str, filters: Filters) -> Dict[str, float]:
	"""Returns a dict of leave type and corresponding leaves taken by employee like:
	{'leave_without_pay': 1.0, 'sick_leave': 2.0}
	"""
	Attendance = frappe.qb.DocType("Attendance")
	day_case = frappe.qb.terms.Case().when(Attendance.status == "Half Day", 0.5).else_(1)
	sum_leave_days = Sum(day_case).as_("leave_days")

	leave_details = (
		frappe.qb.from_(Attendance)
		.select(Attendance.leave_type, sum_leave_days)
		.where(
			(Attendance.employee == employee)
			& (Attendance.docstatus == 1)
			& (Attendance.company == filters.company)
			& ((Attendance.leave_type.isnotnull()) | (Attendance.leave_type != ""))
			& (Extract("month", Attendance.attendance_date) == filters.month)
			& (Extract("year", Attendance.attendance_date) == filters.year)
		)
		.groupby(Attendance.leave_type)
	).run(as_dict=True)

	leaves = {}
	for d in leave_details:
		leave_type = frappe.scrub(d.leave_type)
		leaves[leave_type] = d.leave_days

	return leaves


def get_entry_exits_summary(employee: str, filters: Filters) -> Dict[str, float]:
	"""Returns total late entries and total early exits for employee like:
	{'total_late_entries': 5, 'total_early_exits': 2}
	"""
	Attendance = frappe.qb.DocType("Attendance")

	late_entry_case = frappe.qb.terms.Case().when(Attendance.late_entry == "1", "1")
	count_late_entries = Count(late_entry_case).as_("total_late_entries")

	early_exit_case = frappe.qb.terms.Case().when(Attendance.early_exit == "1", "1")
	count_early_exits = Count(early_exit_case).as_("total_early_exits")

	entry_exits = (
		frappe.qb.from_(Attendance)
		.select(count_late_entries, count_early_exits)
		.where(
			(Attendance.docstatus == 1)
			& (Attendance.employee == employee)
			& (Attendance.company == filters.company)
			& (Extract("month", Attendance.attendance_date) == filters.month)
			& (Extract("year", Attendance.attendance_date) == filters.year)
		)
	).run(as_dict=True)

	return entry_exits[0]


@frappe.whitelist()
def get_attendance_years() -> str:
	"""Returns all the years for which attendance records exist"""
	Attendance = frappe.qb.DocType("Attendance")
	year_list = (
		frappe.qb.from_(Attendance)
		.select(Extract("year", Attendance.attendance_date).as_("year"))
		.distinct()
	).run(as_dict=True)

	if year_list:
		year_list.sort(key=lambda d: d.year, reverse=True)
	else:
		year_list = [frappe._dict({"year": getdate().year})]

	return "\n".join(cstr(entry.year) for entry in year_list)


def get_chart_data(attendance_map: Dict, filters: Filters) -> Dict:
	days = get_columns_for_days(filters)
	labels = []
	absent = []
	present = []
	leave = []

	for day in days:
		labels.append(day["label"])
		total_absent_on_day = total_leaves_on_day = total_present_on_day = 0

		for employee, attendance_dict in attendance_map.items():
			for shift, attendance in attendance_dict.items():
				attendance_on_day = attendance.get(cint(day["fieldname"]))

				if attendance_on_day == "On Leave":
					# leave should be counted only once for the entire day
					total_leaves_on_day += 1
					break
				elif attendance_on_day == "Absent":
					total_absent_on_day += 1
				elif attendance_on_day in ["Present", "Work From Home"]:
					total_present_on_day += 1
				elif attendance_on_day == "Half Day":
					total_present_on_day += 0.5
					total_leaves_on_day += 0.5

		absent.append(total_absent_on_day)
		present.append(total_present_on_day)
		leave.append(total_leaves_on_day)

	return {
		"data": {
			"labels": labels,
			"datasets": [
				{"name": "Absent", "values": absent},
				{"name": "Present", "values": present},
				{"name": "Leave", "values": leave},
			],
		},
		"type": "line",
		"colors": ["red", "green", "blue"],
	}
