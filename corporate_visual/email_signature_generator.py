# -*- coding: utf-8 -*-
"""
Created on Tue Oct 31 09:30:57 2017
Updated on Thu Oct 31          2017

@author: Elan H. Herrera
@description:
    This is the template used to construct the HTML code for the Corporate
    Visual Communications email signature. This script will embed individual
    user data in a standardized template to generate uniform, company email
    signatures for each user.
"""

# --------------------------------------------------------------------------- #
import os
from string import Template
from configparser import ConfigParser
from bs4 import BeautifulSoup, Comment
from premailer import transform

# --------------------------------------------------------------------------- #
# Script file definitions
signature_template_file = "signature_template.html"
company_info_file = "info_company.cfg"
employee_info_file = "info_employees.cfg"
output_folder = "signature_files"

# --------------------------------------------------------------------------- #
# DEF : make directory path
def mkdir(path):
    try:
        os.makedirs(path)
    except OSError:
        if not os.path.isdir(path):
            raise

# --------------------------------------------------------------------------- #
# DEF : remove optional mobile phone row
def remove_mobile_phone(soup):
    for row in soup('tr', {'class':'optional_phone'}):
        row.extract()

# --------------------------------------------------------------------------- #
# DEF : remove all comment lines
def remove_comments(soup):
    for element in soup(text=lambda text: isinstance(text, Comment)):
        element.extract()

# --------------------------------------------------------------------------- #
# DEF : remove any new line returns
def remove_new_lines(str_in):
    str_out = str_in.replace('\r', '').replace('\n', '')
    return str_out

# --------------------------------------------------------------------------- #
# DEF: remove all excess white spaces
def remove_white_spaces(str_in):
    str_out = ' '.join(str_in.split())
    return str_out

# --------------------------------------------------------------------------- #
# DEF : generate the template from soup
def build_template(soup):
    remove_comments(soup)
    html_string = str(soup)
#    html_inline = remove_new_lines(html_string)
#    html_tight = remove_white_spaces(html_inline)
    html_template = Template(html_string)
    return(html_template)

# --------------------------------------------------------------------------- #
# Create output file directory
if not os.path.isdir(output_folder):
    mkdir(output_folder)
else:
    for content in os.listdir(output_folder):
        path = os.path.join(output_folder, content)
        try:
            if os.path.isfile(path):
                os.unlink(path)
        except:
            print("exception on '%s'!" % path)

# --------------------------------------------------------------------------- #
# Read the configuration files containing company and employee information
company = ConfigParser(allow_no_value=True)
company.read(company_info_file)

employees = ConfigParser(allow_no_value=True)
employees.read(employee_info_file)

# --------------------------------------------------------------------------- #
# Set the basic attribues of the signature
data = {} # generates a dictionary object for key/value pairs

data['contact_icon_size'] = 20 # all link icons are square and uniform size
data['link_icon_size'] = 24 # all link icons are square and uniform size

# --------------------------------------------------------------------------- #
# Configure company information variables
address_1 = company.get('Address','number')
address_1 = address_1 + " " + company.get('Address','street')

if company.has_option('Address','suite_number'):
    address_1 = address_1 + ", " + company.get('Address','suite_type')
    address_1 = address_1 + " " + company.get('Address','suite_number')

address_2 = company.get('Address','city')
address_2 = address_2 + ", " + company.get('Address','state')
address_2 = address_2 + " " + company.get('Address','zipcode')

# --------------------------------------------------------------------------- #
# Substitute company information variables
data['company_name']        = company.get('General','name_full')
data['company_initials']    = company.get('General','name_initials')

# --------------------------------------------------------------------------- #
data['address_line_1']      = address_1
data['address_line_2']      = address_2

# --------------------------------------------------------------------------- #
data['phone_number_office'] = company.get('Phone','office')
data['phone_number_fax']    = company.get('Phone','fax')

# --------------------------------------------------------------------------- #
data['phone_icon_office']   = company.get('PhoneImage','office')
data['phone_icon_fax']      = company.get('PhoneImage','fax')
data['phone_icon_mobile']   = company.get('PhoneImage','mobile')

# --------------------------------------------------------------------------- #
data['link_site_0']         = company.get('LinkSite','company')
data['link_icon_0']         = company.get('LinkImage','company')

# --------------------------------------------------------------------------- #
data['link_name_1']         = 'Home Page'
data['link_site_1']         = company.get('LinkSite','company')
data['link_icon_1']         = company.get('LinkImage','website')
data['link_over_1']         = company.get('LinkMessege','company')

# --------------------------------------------------------------------------- #
data['link_name_2']         = 'Google'
data['link_site_2']         = company.get('LinkSite','google')
data['link_icon_2']         = company.get('LinkImage','google')
data['link_over_2']         = company.get('LinkMessege','google')

# --------------------------------------------------------------------------- #
data['link_name_3']         = 'Facebook'
data['link_site_3']         = company.get('LinkSite','facebook')
data['link_icon_3']         = company.get('LinkImage','facebook')
data['link_over_3']         = company.get('LinkMessege','facebook')

# --------------------------------------------------------------------------- #
data['link_name_4']         = 'LinkedIn'
data['link_site_4']         = company.get('LinkSite','linkedin')
data['link_icon_4']         = company.get('LinkImage','linkedin')
data['link_over_4']         = company.get('LinkMessege','linkedin')

# --------------------------------------------------------------------------- #
data['link_name_5']         = 'Dropbox Uploads'
# dropbox upload site link set below per employee
data['link_icon_5']         = company.get('LinkImage','dropbox')
data['link_over_5']         = company.get('LinkMessege','dropbox')

# --------------------------------------------------------------------------- #
data['conf_notice']         = company.get('Other','conf_notice')

# --------------------------------------------------------------------------- #
# Configure employee information variables
for person in employees.sections():
    print("=> Loading employee : " + person + "...\n")

    # Make the soup for current employee
    with open(signature_template_file) as file_in:
       soup = BeautifulSoup(file_in, 'lxml')

    # Check if employee or general email
    if not employees.has_option(person,'name_first'):
        employee_name = employees.get(person,'title')
        employee_separator = ""
        employee_position = ""
    else:
        employee_name = employees.get(person,'name_first') + " "
        employee_separator = "|"
        employee_position = employees.get(person,'title')

    # Check if employee has middle name
    if employees.has_option(person,'name_middle'):
        middle_name = employees.get(person,'name_middle')
        middle_init = middle_name[0]
        employee_name = employee_name + middle_init + ". "

    # Check for last name or generic role
    if employees.has_option(person,'name_last'):
        last_name = employees.get(person,'name_last')
        employee_name = employee_name + last_name

    # Check if employee has suffix (professional title)
    if employees.has_option(person,'name_suffix'):
        employee_suffix = employees.get(person,'name_suffix')
        employee_name = employee_name + ", " + employee_suffix

    # Check for mobile phone number
    if employees.has_option(person,'mobile'):
        employee_phone_number = employees.get(person,'mobile')
    else:
        employee_phone_number = "test"
        remove_mobile_phone(soup)

    # Check for personal Dropobox upload link
    if employees.has_option(person,'dropbox'):
            data['link_site_5'] = employees.get(person,'dropbox')
    else:
            data['link_site_5'] = company.get('LinkSite','dropbox')

# --------------------------------------------------------------------------- #
# Substitute employee information variables
    data['employee_name'] = employee_name
    data['employee_separator'] = employee_separator
    data['employee_position'] = employee_position
    data['phone_number_mobile'] = employee_phone_number

# --------------------------------------------------------------------------- #
# Generate the output HTML
    # Substitute data into HTML template
    html_src = build_template(soup)
    html_filled = html_src.substitute(data)

    # Convert CSS HTML to inline HTML
    html_out = transform(html_filled)
    html_out = remove_new_lines(html_out)
    html_out = remove_white_spaces(html_out)

    # Write HTML code to file
    output_name = "signature_" + person + ".html"
    output_path = os.path.join(output_folder, output_name)
    with open(output_path, "w") as file_out:
        file_out.write(html_out)

# --------------------------------------------------------------------------- #
# Print process completion messege
print("... Employee signatures were successfully generated!")
