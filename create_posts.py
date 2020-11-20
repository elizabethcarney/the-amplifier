"""
scrape_sheets.py
Elizabeth Carney
Created for The Amplifier 2020
"""

from __future__ import print_function
import string
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# spreadsheet: Combed Data Participant Responses
SPREADSHEET_ID = '1Fycidn_h-GyS9Wnl2uVsE9wlY4zG0jDsCpO0WEuh0J0'
RANGE = 'Bios!A2:R35'

# if modifying these scopes, delete the file token.pickle
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def oauth():
    """
    Authorizes API calls
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds

def get_email_prefix(email):
    """
    returns substring before @ symbol in email address
    """
    i_at = email.find('@')
    if i_at != -1:
        prefix = email[:i_at]
        return prefix

def find_nicename(name):
    """
    finds project's nicename from its title
    """
    nicename = name.lower()
    trans_table = nicename.maketrans(' ','-',string.punctuation)
    nicename = nicename.translate(trans_table)
    return nicename

def build_content(row):
    """
    returns string of post content based on form values

    NOTE: posts currently will not have images!
    """
    # name row values for readability
    name = row[2]
    pron = row[3]
    bio = row[4]
    site = row[6]
    genrole = row[7]

    # name_str
    if pron == "":
        name_str = '<b>' + name + '</b>\n\n'
    else:
        name_str = '<b>' + name + ', ' + pron + '</b>\n\n'
    # roles_str
    roles_str = genrole
    if (len(row) > 9) and (row[8] != ""):
        proj_nicename = find_nicename(row[8])
        roles_str += '<a href="https://sophia.smith.edu/theamplifier/portfolio/' + proj_nicename + '/" /><em>' + row[8] + '</em></a> (' + row[9] + ')'
    for i in range(10, len(row), 2):
        if row[i] != "":
            proj_nicename = find_nicename(row[i])
            roles_str += ', <a href="https://sophia.smith.edu/theamplifier/portfolio/' + proj_nicename + '/" /><em>' + row[i] + '</em></a> (' + row[i+1] + ')'
    # site_str
    site_str = ""
    if site != "":
        site_str = '\n\n<a href="' + site + '" />' + site + '</a>'

    content_str = name_str + bio + '\n\nRoles: ' + roles_str + site_str
    return content_str

def get_categories(row):
    """
    Gets list of person's projects with nicenames
    """
    categories = []

    for i in range(8, len(row), 2):
        if row[i] != "":
            proj = row[i]
            proj_nicename = find_nicename(proj)
            categories.append({ "name": proj, "nicename": proj_nicename })

    return categories

def main():
    """
    Creates XML file to be imported into Wordpress autofilled with responses from Google spreadsheet
    """

    # initialize API service
    creds = oauth()
    service = build('sheets', 'v4', credentials=creds)

    # call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
    values = result.get('values', [])

    # initialize variables
    post_title = ""
    post_name = ""
    post_content = ""
    post_id = 201
    post_categories = []

    # static parts of xml file to write
    xmlrss_open = '<?xml version="1.0" encoding="UTF-8" ?>\n<rss version="2.0"\nxmlns:excerpt="http://wordpress.org/export/1.2/excerpt/"\nxmlns:content="http://purl.org/rss/1.0/modules/content/"\nxmlns:wfw="http://wellformedweb.org/CommentAPI/"\nxmlns:dc="http://purl.org/dc/elements/1.1/"\nxmlns:wp="http://wordpress.org/export/1.2/"\n>\n'
    channel_open = '<channel>\n<title>The Amplifier Project</title>\n<link>https://sophia.smith.edu/theamplifier</link>\n<description>Presented by the Smith College Department of Theatre</description>\n<pubDate>Thu, 19 Nov 2020 10:00:00 +0000</pubDate>\n<language>en-US</language>\n<wp:wxr_version>1.2</wp:wxr_version>\n<wp:base_site_url>http://sophia.smith.edu/</wp:base_site_url>\n<wp:base_blog_url>https://sophia.smith.edu/theamplifier</wp:base_blog_url>\n'
    author_open = '<wp:author><wp:author_id>1502</wp:author_id><wp:author_login><![CDATA[ekcarney@smith.edu]]></wp:author_login><wp:author_email><![CDATA[ekcarney@smith.edu]]></wp:author_email><wp:author_display_name><![CDATA[ekcarney@smith.edu]]></wp:author_display_name><wp:author_first_name><![CDATA[Elizabeth]]></wp:author_first_name><wp:author_last_name><![CDATA[Carney]]></wp:author_last_name></wp:author>\n<generator>https://wordpress.org/?v=5.2.2</generator>\n'
    static_opener = xmlrss_open + channel_open + author_open
    static_created = '<pubDate>Fri, 13 Nov 2020 10:00:00 +0000</pubDate>\n<dc:creator><![CDATA[ekcarney@smith.edu]]></dc:creator>\n'
    static_desc = '<description></description>\n'
    static_encoded = '<excerpt:encoded><![CDATA[]]></excerpt:encoded>\n'
    static_datetime = '<wp:post_date><![CDATA[2020-11-13 10:00:00]]></wp:post_date>\n<wp:post_date_gmt><![CDATA[2020-11-13 15:00:00]]></wp:post_date_gmt>\n<wp:comment_status><![CDATA[closed]]></wp:comment_status>\n<wp:ping_status><![CDATA[open]]></wp:ping_status>\n'
    static_relations = '<wp:status><![CDATA[publish]]></wp:status>\n<wp:post_parent>0</wp:post_parent>\n<wp:menu_order>0</wp:menu_order>\n<wp:post_type><![CDATA[post]]></wp:post_type>\n<wp:post_password><![CDATA[]]></wp:post_password>\n<wp:is_sticky>0</wp:is_sticky>\n'
    static_postmeta = '<wp:postmeta>\n<wp:meta_key><![CDATA[_edit_last]]></wp:meta_key>\n<wp:meta_value><![CDATA[3977]]></wp:meta_value>\n</wp:postmeta>\n<wp:postmeta>\n<wp:meta_key><![CDATA[_wp_page_template]]></wp:meta_key>\n<wp:meta_value><![CDATA[default]]></wp:meta_value>\n</wp:postmeta>\n'
    static_closer = '</channel>\n</rss>\n'

    if not values:
        print('No data found.')
    else:
        # create xml file using spreadsheet data
        filename = 'posts.xml'
        f = open(filename, "a")
        # write xml header
        f.write(static_opener)

        # TODO
        #for row in values:
        for i in range(0,2,1):
            row = values[i]

            # name row values for readability
            form_email = row[0]
            form_name = row[2]

            # set details for this post
            post_title = form_name
            post_name = get_email_prefix(form_email)
            post_content = build_content(row)
            post_categories = get_categories(row)

            # build post xml
            # variable tags: title, link, guid, content-encoded, wp:post_id, wp:post_name, category
            f.write('<item>\n<title>' + post_title + '</title>\n')
            f.write('<link>https://sophia.smith.edu/theamplifier/' + post_name + '/</link>\n')
            f.write(static_created)
            f.write('<guid isPermaLink="false">https://sophia.smith.edu/theamplifier/?p=' + str(post_id) + '</guid>\n')
            f.write(static_desc)
            f.write('<content:encoded><![CDATA[' + post_content + ']]>\n</content:encoded>\n')
            f.write(static_encoded)
            f.write('<wp:post_id>' + str(post_id) + '</wp:post_id>\n')
            f.write(static_datetime)
            f.write('<wp:post_name><![CDATA[' + post_name + ']]></wp:post_name>\n')
            f.write(static_relations)
            for cat in post_categories:
                f.write('<category domain="category" nicename="' + cat['nicename'] + '"><![CDATA[' + cat['name'] + ']]></category>\n')
            f.write(static_postmeta)
            f.write('</item>\n')

            post_id += 1

        f.write(static_closer)
        f.close()


if __name__ == '__main__':
    main()
