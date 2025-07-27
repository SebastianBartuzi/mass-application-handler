import os
from dotenv import load_dotenv


class BaseConfig:

    load_dotenv()

    project_id = os.getenv('PROJECT_ID')
    location_code = os.getenv('LOCATION_CODE')
    response_model = os.getenv('RESPONSE_MODEL')
    thinking_budget = int(os.getenv('THINKING_BUDGET'))

    sender_email = os.getenv('SENDER_EMAIL')
    confirmation_addressees = os.getenv('CONFIRMATION_ADDRESSEES')
    
    excel_file_start_row = int(os.getenv('EXCEL_FILE_START_ROW'))
    excel_file_teryt_column = os.getenv('EXCEL_FILE_TERYT_COLUMN')
    excel_file_authority_name_column = os.getenv('EXCEL_FILE_AUTHORITY_NAME_COLUMN')
    excel_file_governor_title_column = os.getenv('EXCEL_FILE_GOVERNOR_TITLE_COLUMN')
    excel_file_governor_gender_column = os.getenv('EXCEL_FILE_GOVERNOR_GENDER_COLUMN')
    excel_file_governor_name_column = os.getenv('EXCEL_FILE_GOVERNOR_NAME_COLUMN')
    excel_file_authority_office_name_column = os.getenv('EXCEL_FILE_AUTHORITY_OFFICE_NAME_COLUMN')
    excel_file_authority_office_mail_column = os.getenv('EXCEL_FILE_AUTHORITY_OFFICE_MAIL_COLUMN')

    excel_file_path = os.getenv('EXCEL_FILE_PATH')
    application_template_path = os.getenv('APPLICATION_TEMPLATE_PATH')
    mail_analysis_prompt_path = os.getenv('MAIL_ANALYSIS_PROMPT_PATH')
    mail_analysis_response_schema_path = os.getenv('MAIL_ANALYSIS_PROMPT_PATH')
    teryt_matcher_prompt_path = os.getenv('TERYT_MATCHER_RESPONSE_SCHEMA_PATH')
    teryt_matcher_response_schema_path = os.getenv('TERYT_MATCHER_RESPONSE_SCHEMA_PATH')

    temp_path = os.getenv('TMP_PATH')
    attachments_dir_path = os.getenv('ATTACHMENTS_DIR_PATH')