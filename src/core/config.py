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
    
    data_excel_tab_name = os.getenv('DATA_EXCEL_TAB_NAME')
    data_excel_start_row = int(os.getenv('DATA_EXCEL_START_ROW'))
    data_excel_teryt_column = os.getenv('DATA_EXCEL_TERYT_COLUMN')
    data_excel_authority_name_column = os.getenv('DATA_EXCEL_AUTHORITY_NAME_COLUMN')
    data_excel_governor_title_column = os.getenv('DATA_EXCEL_GOVERNOR_TITLE_COLUMN')
    data_excel_governor_gender_column = os.getenv('DATA_EXCEL_GOVERNOR_GENDER_COLUMN')
    data_excel_governor_name_column = os.getenv('DATA_EXCEL_GOVERNOR_NAME_COLUMN')
    data_excel_authority_office_name_column = os.getenv('DATA_EXCEL_AUTHORITY_OFFICE_NAME_COLUMN')
    data_excel_authority_office_mail_column = os.getenv('DATA_EXCEL_AUTHORITY_OFFICE_MAIL_COLUMN')

    report_excel_tab_name = os.getenv('REPORT_EXCEL_TAB_NAME')
    report_excel_start_row = int(os.getenv('REPORT_EXCEL_START_ROW'))
    report_excel_teryt_column = os.getenv('REPORT_EXCEL_TERYT_COLUMN')
    report_excel_last_answer_date_column = os.getenv('REPORT_EXCEL_LAST_ANSWER_DATE_COLUMN')
    report_excel_needs_attention_column = os.getenv('REPORT_EXCEL_NEEDS_ATTENTION_COLUMN')
    report_excel_additional_context_column = os.getenv('REPORT_EXCEL_ADDITIONAL_CONTEXT_COLUMN')
    report_excel_offers_internships_column = os.getenv('REPORT_EXCEL_OFFERS_INTERNSHIPS_COLUMN')
    report_excel_are_internships_paid_column = os.getenv('REPORT_EXCEL_ARE_INTERNSHIPS_PAID_COLUMN')
    report_excel_plans_paid_internships = os.getenv('REPORT_EXCEL_PLANS_PAID_INTERNSHIPS')
    report_excel_paid_internships_number_column = os.getenv('REPORT_EXCEL_PAID_INTERNSHIPS_NUMBER_COLUMN')
    report_excel_internships_salaries_column = os.getenv('REPORT_EXCEL_INTERNSHIPS_SALARIES_COLUMN')
    report_excel_additional_info_column = os.getenv('REPORT_EXCEL_ADDITIONAL_INFO_COLUMN')

    data_excel_file_path = os.getenv('DATA_EXCEL_FILE_PATH')
    report_excel_file_path = os.getenv('REPORT_EXCEL_FILE_PATH')
    report_template_excel_file_path = os.getenv('REPORT_TEMPLATE_EXCEL_FILE_PATH')
    application_template_path = os.getenv('APPLICATION_TEMPLATE_PATH')
    mail_analysis_prompt_path = os.getenv('MAIL_ANALYSIS_PROMPT_PATH')
    mail_analysis_response_schema_path = os.getenv('MAIL_ANALYSIS_PROMPT_PATH')
    teryt_matcher_prompt_path = os.getenv('TERYT_MATCHER_RESPONSE_SCHEMA_PATH')
    teryt_matcher_response_schema_path = os.getenv('TERYT_MATCHER_RESPONSE_SCHEMA_PATH')

    temp_path = os.getenv('TMP_PATH')
    attachments_dir_path = os.getenv('ATTACHMENTS_DIR_PATH')