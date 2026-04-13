import pandas as pd
from datetime import date as dt
from datetime import timedelta
from datetime import datetime
import win32com.client as win32
import mappings
from snapshot_logger import logger
from glob import glob
from decorator import wait_for_file
from tkinter import simpledialog as sd
from tkinter import messagebox as mb

class DateFunctions:
    def __init__(self) -> None:
        self.date = None
        pass

    # Take input date, get dow, change date if it is a weekend, get date in string format
    def get_file_date(self, initial_date):
        logger.info("Getting the current date")
        logger.debug(f"Current date: {initial_date}")
        logger.debug(type(initial_date))
        self.today = initial_date
        self.file_date = self.change_date(days= -1, date= self.today)
        logger.info(f"Current date is {self.file_date}")
        self.file_date_wd = self.file_date.weekday()
        logger.debug(f"Current date weekday is {self.file_date_wd}")
        self.file_date = self.check_date(dow= self.file_date_wd, date= self.file_date)
        self.get_file_date_str(date= self.file_date)
        return self.file_date

    # get the current date in string format
    def get_file_date_str(self, date):
        self.file_date_str = date.strftime("%m%d%Y")
        self.file_date_str_slash = date.strftime("%m/%d/%Y")
        self.file_date_m_str = date.strftime("%m")
        self.file_date_d_str = date.strftime("%d")
        self.file_date_y_str = date.strftime("%Y")

    # change the date by a certain number of days
    def change_date(self, days, date):
        logger.info(f"adding {days} day(s) to {date}")
        new_date = date + timedelta(days=days)
        logger.debug(f"new date is: {new_date}")
        return new_date
    
    # check if the date is a saturday or sunday. if yes, change the date to the previous friday
    def check_date(self, dow, date):
        logger.info("Checking file date is not a weekend")
        if dow == 5:
            logger.debug("file date is a saturday")
            new_date = self.change_date(days= -1, date= date)
        elif dow == 6:
            logger.debug("file date is a sunday")
            new_date = self.change_date(days= -2, date= date)
        else:
            logger.debug("file date is a weekday")
            new_date = date
        return new_date

    # parse the date string into a datetime object. Format MMDDYYYY
    def parse_date(self, date_str):
        return datetime.strptime(date_str, "%m%d%Y")
    
    # ask the user if the date is correct. If not, ask for a new date
    def ask_if_correct_date(self, today):
        answer = mb.askyesno(f"Check Date", f"Do you want to run for the calendar date of {today}?")
        if not answer:
                day = sd.askinteger("Follow Up", "Please enter day of the month as number : ", minvalue=1, maxvalue=31)
                logger.debug(f"{day} was entered")
                month = sd.askinteger("Follow Up", "Please enter a month number (e.g. March = 3): ", minvalue=1, maxvalue=12)
                logger.debug(f"{month} was entered")
                year = sd.askinteger("Follow Up", "Please enter a year: ", minvalue=2020, maxvalue=2030)
                logger.debug(f"{year} was entered")
                try:
                    date = today.replace(day=day, month=month, year=year)
                    logger.info(f"new file date is {date}")
                except TypeError:
                    logger.critical("No date selected")
                    logger.info("Stopping Process")
                    exit()
                else:
                    return date
        else:
            return today
    

class MainSpreadsheet(DateFunctions):
    def __init__(self) -> None:
        super().__init__()
        pass

    # run the methods that will read the main spreadsheet
    def run(self, init_date):
        # self.combined_op_file_path = "M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/GOA/Inputs/Outbound_"
        self.combined_op_file_path = "M:/CPP-Data/Sutherland RPA/Combined Outputs/Outbound_"
        self.file_date = self.get_file_date(initial_date= init_date)
        self.get_file_path()
        self.main_df = self.read_main_spreadsheet(file=self.file_path)
    
    # get the file path for the main spreadsheet
    def get_file_path(self):
        logger.info("Getting the file path")
        self.file_path = self.combined_op_file_path + self.file_date_str + ".xlsx"
    
    # wait for the file to be available
    @wait_for_file()
    def read_main_spreadsheet(self, file): # read the main spreadsheet
        logger.info("Reading the main spreadsheet")
        main_df = pd.read_excel(
            file, 
            engine="openpyxl", 
            na_values=" ", 
            keep_default_na=False
        )
        return main_df


class Snapshot(MainSpreadsheet):
    def __init__(self, use_case, main_df, file_date, mappings_dict) -> None:
        super().__init__()
        self.use_case = use_case
        self.file_path = mappings_dict[f"{self.use_case}"]["file_path"]
        self.status_crosswalk = mappings_dict[f"{self.use_case}"]["status_crosswalk"]
        self.scenario_crosswalk = mappings_dict[f"{self.use_case}"]["scenario_crosswalk"]
        self.columns = mappings_dict[f"{self.use_case}"]["columns"]
        self.columns_crosswalk = mappings_dict[f"{self.use_case}"]["column_crosswalk"]
        self.cc_emails = mappings_dict[f"{self.use_case}"]["carbon_copy"]
        self.name_format_str = mappings_dict[f"{self.use_case}"]["name_format"]
        self.drop_columns = mappings_dict[f"{self.use_case}"]["drop_columns"]
        self.botname = mappings_dict[f"{self.use_case}"]["BotName"]
        self.main_df = main_df
        self.file_date = file_date   
    
    # run the methods that will parse the spreadsheet
    def parse_spreadsheet(self):
        logger.info(f"Parsing the spreadsheet for {self.use_case}")
        self.use_case_df = self.main_df[self.main_df["BotName"] == self.botname]
        if self.use_case == "NCOA":
            self.get_ncoa_date() # Will reset self.file_date + dependents; NCOA must be last in mappings_dict
        self.export_to_excel()
        self.use_case_df = pd.DataFrame(self.use_case_df, columns=self.columns)
        self.use_case_df.rename(columns=self.columns_crosswalk, inplace=True)
        self.use_case_df['RD + Reason'] = self.use_case_df['Retrieval Description']+ " - " + self.use_case_df['Reason']

    # get the date for the NCOA use case
    def get_ncoa_date(self):
        # set new file date by looking at column 'BOTRequestDate' in self.use_case_df
        self.file_date = self.use_case_df['BOTRequestDate'].iloc[0] # get the date from the first row
        # take value in self.file_date input template: 3/26/2024 7:25:15 AM and turn it into a datetime object
        self.file_date = datetime.strptime(self.file_date, "%m/%d/%Y %I:%M:%S %p")
        logger.debug(f"date - {self.file_date} type - {type(self.file_date)}")
        self.get_file_date_str(date= self.file_date)

    # export the dataframe to an excel file
    def export_to_excel(self):
        logger.info(f"Exporting to Excel for {self.use_case}")
        self.get_file_date_str(self.file_date)
        self.get_name_format()
        self.use_case_df.to_excel(self.name_format, index=False)

    # get the name format for the file
    def get_name_format(self):
        logger.info(f"Getting the name format for {self.use_case}")
        self.name_format = self.name_format_str.format(
            file_path= self.file_path,
            month_str= self.file_date_m_str, 
            day_str= self.file_date_d_str, 
            year_str= self.file_date_y_str, 
        )

    # get the business status for each row
    def get_business_status(self, row):
        return self.status_crosswalk.get(row['RD + Reason'], 'Exception')
    
    # get the business scenario for each row
    def get_business_scenario(self, row):
        return self.scenario_crosswalk.get(row['RD + Reason'], 'Unmapped Exception Scenario')
    
    # apply the business rules to the dataframe
    def apply_business_rules(self):
        logger.info(f"Applying business rules for {self.use_case}")
        logger.info(f"Getting business statuses for {self.use_case}")
        self.use_case_df['Business Status'] = self.use_case_df.apply(
            lambda row: self.get_business_status(row), 
            axis=1
            )
        logger.info(f"Getting business scenarios for {self.use_case}")
        self.use_case_df['Business Scenario'] = self.use_case_df.apply(
            lambda row: self.get_business_scenario(row), 
            axis=1
            )
    
    # get the exceptions for the use case
    def get_exceptions(self):
        logger.info(f"Getting exceptions for {self.use_case}")
        self.exceptions_df = pd.DataFrame(self.use_case_df)
        self.exceptions_df = self.use_case_df[self.use_case_df['Business Status'] == 'Exception']
        self.exceptions_df = self.exceptions_df.drop(columns=self.drop_columns, axis=1)
        self.exceptions_df = self.exceptions_df.sort_values(by='Business Scenario', ascending=False)
        self.html_table = self.exceptions_df.to_html(
            index=False, 
            classes="dataframe", 
            border=2, 
            justify="center", 
            table_id="exceptions_table"
            )

    # calculate the results for the use case
    def calc_results(self):
        logger.info(f"Calculating results for {self.use_case}")
        self.total_rows = len(self.use_case_df)
        self.bus_status_counts = self.use_case_df['Business Status'].value_counts()
        self.bus_status_percent = self.bus_status_counts / self.total_rows * 100
        self.bus_status_percent = self.bus_status_percent.round(2)
        self.rd_reason_counts = self.use_case_df['RD + Reason'].value_counts()

        self.get_exceptions()
    
    # write the email body for the use case
    def write_email_body(self):
        logger.info(f"Writing email body for {self.use_case}")
        self.email_body = f"""
        <p><strong>Total Cases Processed:</strong> {self.total_rows}</p>
        <p>Link to File can be found: <a href="file:///{self.name_format}">here</a></p>
        <p><strong>Count for each Description - Reason:</strong></p>
        """
        for index, count in self.rd_reason_counts.items():
                self.email_body += f"<p>{index}: {count}</p>"
        
        self.email_body += f"""
        <strong>Rate for each Business Status:</strong><br>
        Success Rate - {self.bus_status_percent.get('Success', 0.0)}%<br>
        Exception Rate - {self.bus_status_percent.get('Exception', 0.0)}%<br><br>
        <strong>List of Exceptions:</strong>
        {self.html_table}
        """

    # compose and send the email
    def compose_and_send(self):
        self.outlook = win32.Dispatch('Outlook.Application')
        self.mail = self.outlook.CreateItem(0)
        self.mail.Subject = f'{self.use_case} Daily Snapshot - {self.file_date_str_slash}'
        self.mail.HTMLBody = self.email_body
        self.mail.To = 'denglish2@northwell.edu'
        self.mail.CC = self.cc_emails
        self.mail.Send()