import sys
import os

try:
    from jira import JIRA
except ImportError:
    sys.exit("""
You need the JIRA Python library

run
pip install jira
""")
try:
    import xlsxwriter
except ImportError:
    sys.exit("""
You need the 'xlsxwriter' Python library

run
pip install xlswriter
""")
#import requests
#import json

filename=os.path.basename(sys.argv[0])
sys.argv.pop(0)
row=0

if len(sys.argv) == 0:
    print("No issue found...\n")
    print("Provide one or more JIRA issue (Story, Epic, or Initiative)")
    print("  USAGE: %s [INIT-ID1] [INIT-ID2] ...\n"% filename)
    exit(1)

jira = JIRA(options={'server': 'https://jira.domain.com'})

def get_epic_ids ( initiative_id ):
    issue = jira.issue(initiative_id)
    epics = (issue.fields.customfield_19118)
    keys = sorted([epic.key for epic in epics])
    return keys

def get_story_ids ( epic_id ):
    stories = jira.search_issues('"Epic Link" in (' + epic_id + ')')
    keys = sorted([story.key for story in stories])
    return keys

def get_issue_details ( issue_id ):
    issue = jira.issue(issue_id)
    id = issue.key
    status = issue.fields.status.name
    summary = issue.fields.summary
    assigned = ', '.join([assigned.displayName for assigned in issue.fields.customfield_18801] \
        if not(issue.fields.customfield_18801 is None) else '')
    reporter = issue.fields.reporter.displayName or ''
    resolution_date = issue.fields.resolutiondate or ''
    teams = ', '.join([team.value for team in issue.fields.customfield_18915]
        if not(issue.fields.customfield_18915 is None) else '')
    link = "https://jira.uptake.com/browse/%s" % id

    return dict([(k, eval(k)) for k in (
        'id',
        'status',
        'summary',
        'assigned',
        'reporter',
        'resolution_date',
        'teams',
        'link'
        )])

def excel_issue(dict, type, style):
    global row
    sep='\t'
    worksheet.write_row(row, 0, (
        type,
        dict.get('id') if style == 'Initiative' else '',
        dict.get('id') if style == 'Epic' else '',
        dict.get('id') if (style == 'Story' or style == 'Other') else '',
        dict.get('status'),
        #'    '* indent_level +  dict.get('summary'),
        #dict.get('summary'),
        '',
        dict.get('assigned'),
        dict.get('teams'),
        dict.get('reporter'),
        dict.get('resolution_date')
        ))
    #format Summary based on sytle
    if style == 'Initiative':
        worksheet.write(row,5,dict.get('summary'),format_initiative)
    elif style == 'Epic':
        worksheet.write(row,5,dict.get('summary'),format_epic)
    elif style == 'Story':
        worksheet.write(row,5,dict.get('summary'),format_story)
    else:
        worksheet.write(row,5,dict.get('summary'),format_other)
    #insert link with id as the text
    worksheet.write_url(row,10,dict.get('link'),string=dict.get('id'))

    row += 1

def print_issue(dict, type, indent_level):
    sep='|'
    print sep.join([
        type,
        sep*indent_level + dict.get('id') + sep*(2-indent_level),
        dict.get('status'),
        '    '* indent_level +  dict.get('summary'),
        dict.get('assigned'),
        dict.get('reporter'),
        dict.get('resolution_date'),
        dict.get('teams'),
        dict.get('link')
        ])

def crawl_initiative ( initiative_id ):
    details = get_issue_details(initiative_id)
    #print_issue(details,'Initiative',0)
    excel_issue(details,'Initiative','Initiative')
    for epic_id in get_epic_ids(initiative_id):
        crawl_epic(epic_id)

def crawl_epic ( epic_id ):
    details = get_issue_details(epic_id)
    #print_issue(details,'Epic',1)
    excel_issue(details,'Epic','Epic')
    for story_id in get_story_ids(epic_id):
        crawl_story(story_id)

def crawl_story ( story_id ):
    details = get_issue_details(story_id)
    #print_issue(details,'Story',2)
    excel_issue(details,'Story','Story')

#------MAIN----------
workbook = xlsxwriter.Workbook('project.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
format_header = workbook.add_format({'bg_color': 'D9D9D9', 'bold': 1})
format_initiative = workbook.add_format({'bold': 1, 'underline': 1,'text_wrap': True, 'align': 'top'})
format_epic = workbook.add_format({'bold': 1,'indent': 1,'text_wrap': True, 'align': 'top'})
format_story = workbook.add_format({'indent': 2,'text_wrap': True, 'align': 'top'})
format_other = workbook.add_format({'text_wrap': True, 'align': 'top'})
for issue_id in (sys.argv):
    if row > 0:
        row +=1
    worksheet.write_row(row,0,('TYPE','INIT_ID','EPIC_ID','ISSUE_ID','STATUS',
        'SUMMARY','ASSIGNED_TO','ASSIGNED_TEAM','RECORDED_BY','RESOLVED_DATE','LINK'),format_header)
    row += 1
    #what type is it?
    issue_type = jira.issue(issue_id).fields.issuetype.name

    if issue_type == 'Initiative':
        #process initiative
        crawl_initiative(issue_id)
    elif issue_type == 'Epic':
        #process epic
        t = 1
    elif issue_type == 'Story':
        #crawl story
        t=1
    else:
        excel_issue(get_issue_details(issue_id),issue_type,'Other')

format_wrap = workbook.add_format({'text_wrap': True, 'align': 'top'})

worksheet.set_column('B:D',10, format_wrap)
worksheet.set_column('E:E',12, format_wrap)
worksheet.set_column('F:F',55, format_wrap)
worksheet.set_column('G:I',20, format_wrap)
worksheet.set_column('J:J',30, format_wrap)
worksheet.set_column('K:K',35, format_wrap)
#worksheet.conditional_format("F1:F%s"%row, {'type': 'formula', 'criteria': '=IF($A1="Initiative",TRUE,FALSE)', 'format': format_initiative})
#worksheet.conditional_format("F1:F%s"%row, {'type': 'formula', 'criteria': '=IF($A1="Epic",TRUE,FALSE)', 'format': format_epic})
#worksheet.conditional_format("F1:F%s"%row, {'type': 'formula', 'criteria': '=IF($A1="Story",TRUE,FALSE)', 'format': format_story})

workbook.close()
