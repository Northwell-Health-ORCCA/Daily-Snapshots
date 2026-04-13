from mappings import mappings_dict
from snapshot_logger import logger
from snapshot_updated import *
from datetime import date as dt

if __name__ == '__main__':
    today = dt.today()
    dt_func = DateFunctions()
    # # date = dt_func.ask_if_correct_date(today)
    date = today
    # date = dt(2026, 3, 20) # For testing purposes only Should be the day after you want to have the snapshot for
    try:
        outbound_df = MainSpreadsheet()
        outbound_df.run(date)
    except FileNotFoundError:
        logger.critical(f"Main outbound spreadsheet not found in M:\CPP-Data\Sutherland RPA\Combined Outputs")
    else:
        for use_case in mappings_dict:
        # for use_case in ['BD IS Printing']:
            snapshot = Snapshot(use_case, outbound_df.main_df, outbound_df.file_date)
            try:
                snapshot.parse_spreadsheet()
                if len(snapshot.use_case_df) == 0:
                    logger.critical(f"No rows in spreadsheet for {use_case}")
                    continue
                snapshot.apply_business_rules()
                snapshot.get_exceptions()
            except AttributeError as e:
                logger.exception(e)
                continue
            except IndexError as e:
                logger.critical(f"No rows in spreadsheet for {use_case}")
                continue
            else:
                snapshot.calc_results()
                snapshot.write_email_body()
                snapshot.compose_and_send()
            
        