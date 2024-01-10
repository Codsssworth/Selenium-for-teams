#!/usr/bin/env python3
import json
import random
import re
import time
import socket
import os
import requests
import selenium
from datetime import datetime
from threading import Timer
import pygetwindow as gw
import pyautogui
from tkinter import *
from ttkthemes import themed_tk as th
from tkterminal import Terminal

from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.utils import ChromeType
from webdriver_manager.microsoft import EdgeChromiumDriverManager

global recconnect
pyautogui.FAILSAFE = False
browser: webdriver.Chrome = None
config = None
meetings = []
current_meeting = None
already_joined_ids = []
active_correlation_id = ""
hangup_thread: Timer = None
conversation_link = "https://teams.microsoft.com/_#/conversations/a"
mode = 3
uuid_regex = r"\b[0-9a-f]{8}\b-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-\b[0-9a-f]{12}\b"
recon = False


class Team:
    def __init__(self, name, t_id, channels=None):
        self.name = name
        self.t_id = t_id
        if channels is None:
            self.get_channels()
        else:
            self.channels = channels

        self.check_blacklist()

    def __str__(self):
        channel_string = '\n\t'.join( [str( channel ) for channel in self.channels] )

        return f"{self.name}\n\t{channel_string}"

    def get_elem(self):
        team_header = browser.find_element_by_css_selector( f"h3[id='{self.t_id}'" )
        team_elem = team_header.find_element_by_xpath( ".." )
        return team_elem

    def expand_channels(self):
        try:
            self.get_elem().find_element_by_css_selector( "div.channels" )
        except exceptions.NoSuchElementException:
            try:
                self.get_elem().click()
                self.get_elem().find_element_by_css_selector( "div.channels" )
            except (exceptions.NoSuchElementException, exceptions.ElementNotInteractableException):
                return None

    def get_channels(self):
        self.expand_channels()
        channels = self.get_elem().find_elements_by_css_selector( ".channels > ul > ng-include > li" )

        channel_names = [channel.get_attribute( "data-tid" ) for channel in channels]
        channel_names = [channel_name[channel_name.find( "channel-" ) + 8:channel_name.find( "-li" )] for channel_name
                         in channel_names if channel_name is not None]

        channels_ids = [channel.get_attribute( "id" ).replace( "channel-", "" ) for channel in channels]

        meeting_states = []
        for channel in channels:
            try:
                channel.find_element_by_css_selector( "a > active-calls-counter" )
                meeting_states.append( True )
            except exceptions.NoSuchElementException:
                meeting_states.append( False )

        self.channels = [Channel( channel_names[i], channels_ids[i], has_meeting=meeting_states[i] ) for i in
                         range( len( channel_names ) )]

    def check_blacklist(self):
        blacklist = config['blacklist']
        blacklist_item = next( (bl_team for bl_team in blacklist if bl_team['team_name'] == self.name), None )
        if blacklist_item is None:
            return

        if len( blacklist_item['channel_names'] ) == 0:
            for channel in self.channels:
                channel.blacklisted = True
        else:
            for channel in self.channels:
                if channel.name in blacklist_item['channel_names']:
                    channel.blacklisted = True


class Channel:
    def __init__(self, name, c_id, blacklisted=False, has_meeting=False):
        self.name = name
        self.c_id = c_id
        self.blacklisted = blacklisted
        self.has_meeting = has_meeting

    def __str__(self):
        return self.name + (" [BLACKLISTED]" if self.blacklisted else "") + (" [MEETING]" if self.has_meeting else "")


class Meeting:
    def __init__(self, m_id, time_started, title, calendar_meeting=False):
        self.m_id = m_id
        self.time_started = time_started
        self.title = title
        self.calendar_meeting = calendar_meeting

    def __str__(self):
        return f"\t{self.title} {self.time_started}" + (" [Calendar]" if self.calendar_meeting else " [Channel]")


def load_config():
    global config
    with open( 'config.json' ) as json_data_file:
        config = json.load( json_data_file )


def init_browser():
    global browser

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument( '--ignore-certificate-errors' )
    chrome_options.add_argument( '--ignore-ssl-errors' )
    chrome_options.add_argument( '--use-fake-ui-for-media-stream' )
    chrome_options.add_experimental_option( 'prefs', {
        'credentials_enable_service': False,
        'profile.default_content_setting_values.media_stream_mic': 1,
        'profile.default_content_setting_values.media_stream_camera': 1,
        'profile.default_content_setting_values.geolocation': 1,
        'profile.default_content_setting_values.notifications': 1,
        'profile': {
            'password_manager_enabled': False
        }
    } )

    chrome_options.add_experimental_option( 'excludeSwitches', ['enable-automation'] )

    if 'headless' in config and config['headless']:
        chrome_options.add_argument( '--headless' )
        print( "Enabled headless mode" )

    if 'mute_audio' in config and config['mute_audio']:
        chrome_options.add_argument( "--mute-audio" )

    if 'chrome_type' in config:
        if config['chrome_type'] == "chromium":
            browser = webdriver.Chrome( ChromeDriverManager( chrome_type=ChromeType.CHROMIUM ).install(),
                                        options=chrome_options )
        elif config['chrome_type'] == "msedge":
            browser = webdriver.Edge( EdgeChromiumDriverManager().install() )
        else:
            browser = webdriver.Chrome( ChromeDriverManager().install(), options=chrome_options )
    else:
        browser = webdriver.Chrome( ChromeDriverManager().install(), options=chrome_options )

    # make the window a minimum width to show the meetings menu
    window_size = browser.get_window_size()
    if window_size['width'] < 1200:
        print( "Resized window" )
        browser.set_window_size( 1200, window_size['height'] )


def wait_until_found(sel, timeout, print_error=True):
    try:
        element_present = EC.visibility_of_element_located( (By.CSS_SELECTOR, sel) )
        WebDriverWait( browser, timeout ).until( element_present )

        return browser.find_element_by_css_selector( sel )
    except exceptions.TimeoutException:
        if print_error:
            print( f"Timeout waiting for element: {sel}" )
        return None


def switch_to_teams_tab():
    teams_button = wait_until_found( "button.app-bar-link > ng-include > svg.icons-teams", 5 )
    if teams_button is not None:
        teams_button.click()


def switch_to_calendar_tab():
    calendar_button = wait_until_found("button.app-bar-link > ng-include > svg.icons-calendar", 5)
    if calendar_button is not None:
        calendar_button.click()

def change_organisation(org_num):
    select_change_org = wait_until_found( "button.tenant-switcher", 20 )
    if select_change_org is None:
        print( "Something went wrong while changing the organisation" )
        return

    select_change_org.click()

    change_org = wait_until_found( f"li.tenant-option[aria-posinset='{org_num}']", 20 )
    if change_org is None:
        print( "Something went wrong while changing the organisation" )
        return

    change_org.click()
    time.sleep( 5 )

    use_web_instead = wait_until_found( ".use-app-lnk", 5, print_error=False )
    if use_web_instead is not None:
        use_web_instead.click()

    time.sleep( 1 )

#
# def prepare_page(include_calendar):
#     try:
#         browser.execute_script( "document.getElementById('toast-container').remove()" )
#     except exceptions.JavascriptException:
#         pass
#
#     if include_calendar:
#         switch_to_calendar_tab()
#
#         view_switcher = wait_until_found( ".ms-CommandBar-secondaryCommand > div > button[class*='__topBarContent']",
#                                           20 )
#         if view_switcher is not None:
#             try:
#                 browser.execute_script( "arguments[0].click();", view_switcher )
#                 time.sleep( 2 )
#             except Exception as e:
#                 print( e )
#                 return
#
#             day_button = wait_until_found(
#                 "li[role='presentation'].ms-ContextualMenu-item>button[aria-posinset='1']", 2, print_error=False )
#             if day_button is None:
#                 browser.execute_script( "arguments[0].click();", view_switcher )
#                 time.sleep( 2 )
#
#             day_button = wait_until_found(
#                 "li[role='presentation'].ms-ContextualMenu-item>button[aria-posinset='1']", 2 )
#             if day_button is not None:
#                 try:
#                     day_button.click()
#                     time.sleep( 2 )
#                 except Exception as e:
#                     print( e )
#                     pass
def prepare_page(include_calendar):
    try:
        browser.execute_script("document.getElementById('toast-container').remove()")
    except exceptions.JavascriptException:
        pass

    if include_calendar:
        switch_to_calendar_tab()

        view_switcher = wait_until_found(".ms-CommandBar-secondaryCommand > div > button[class*='__topBarContent']", 5)

        if view_switcher is not None:
            try:
                browser.execute_script("arguments[0].click();", view_switcher)
                time.sleep(2)
            except Exception as e:
                print(e)
                return

            day_button = wait_until_found(
                "li[role='presentation'].ms-ContextualMenu-item>button[aria-posinset='1']", 2, print_error=False)
            if day_button is None:
                browser.execute_script("arguments[0].click();", view_switcher)
                time.sleep(2)

            day_button = wait_until_found(
                "li[role='presentation'].ms-ContextualMenu-item>button[aria-posinset='1']", 2)
            if day_button is not None:
                try:
                    day_button.click()
                    time.sleep(2)
                except Exception as e:
                    print(e)
                    pass


def get_all_teams():
    team_elems = browser.find_elements_by_css_selector(
        "ul>li[role='treeitem']>div[sv-element]" )

    team_names = [team_elem.get_attribute( "data-tid" ) for team_elem in team_elems]
    team_names = [team_name[team_name.find( 'team-' ) + 5:team_name.rfind( "-li" )] for team_name in team_names]

    team_headers = [team_elem.find_element_by_css_selector( "h3" ) for team_elem in team_elems]
    team_ids = [team_header.get_attribute( "id" ) for team_header in team_headers]

    return [Team( team_names[i], team_ids[i] ) for i in range( len( team_elems ) )]


def get_meetings(teams):
    global meetings

    for team in teams:
        for channel in team.channels:
            if channel.has_meeting and not channel.blacklisted:
                browser.execute_script(
                    f'window.location = "{conversation_link}?threadId={channel.c_id}&ctx=channel";' )

                meeting_elem = wait_until_found( ".ts-calling-thread-header", 10 )
                if meeting_elem is None:
                    continue

                meeting_elems = browser.find_elements_by_css_selector( ".ts-calling-thread-header" )
                for meeting_elem in meeting_elems:
                    meeting_id = meeting_elem.get_attribute( "id" )
                    time_started = int( meeting_id.replace( "m", "" )[:-3] )

                    # already joined calendar meeting
                    correlation_id = meeting_elem.find_element_by_css_selector(
                        "calling-join-button > button" ).get_attribute( "track-data" )
                    if active_correlation_id != "" and correlation_id.find( active_correlation_id ) != -1:
                        continue

                    meetings.append( Meeting( meeting_id, time_started, f"{team.name} -> {channel.name}" ) )


# def get_calendar_meetings():
#     global meetings
#
#     if wait_until_found( "div[class*='__cardHolder']", 20 ) is None:
#         return
#
#     join_buttons = browser.find_elements_by_css_selector(
#         "button[class*='__joinButton'], button[class*='__activeCall']" )
#     if len( join_buttons ) == 0:
#         return
#
#     meeting_cards = []
#     for join_button in join_buttons:
#         meeting_card = join_button.find_element_by_xpath( "../../.." )
#         meeting_cards.append( meeting_card )
#
#     for meeting_card in meeting_cards:
#         style_string = meeting_card.get_attribute( "style" )
#         top_offset = float( style_string[style_string.find( "top: " ) + 5:style_string.find( "rem;" )] )
#
#         minutes_from_midnight = int( top_offset / .135 )
#
#         midnight = datetime.now().replace( hour=0, minute=0, second=0 )
#         midnight = int( datetime.timestamp( midnight ) )
#
#         start_time = midnight + minutes_from_midnight * 60
#
#         sec_meeting_card = meeting_card.find_element_by_css_selector( "div" )
#         meeting_name = sec_meeting_card.get_attribute( "title" ).replace( "\n", " " )
#
#         meeting_id = sec_meeting_card.get_attribute( "id" )
#
#         meetings.append( Meeting( meeting_id, start_time, meeting_name, calendar_meeting=True ) )
def get_calendar_meetings():
    global meetings

    if wait_until_found("div[class*='__cardHolder']", 20, False) is None:
        print("No calendar element found, switch to mode 2 to only check channel meetings")
        return

    join_buttons = browser.find_elements_by_css_selector("button[class*='__joinButton'], button[class*='__activeCall']")
    if len(join_buttons) == 0:
        return

    meeting_cards = []
    for join_button in join_buttons:
        meeting_card = join_button.find_element_by_xpath("../../..")
        meeting_cards.append(meeting_card)

    for meeting_card in meeting_cards:
        style_string = meeting_card.get_attribute("style")
        top_offset = float(style_string[style_string.find("top: ") + 5:style_string.find("rem;")])

        minutes_from_midnight = int(top_offset / .135)

        midnight = datetime.now().replace(hour=0, minute=0, second=0)
        midnight = int(datetime.timestamp(midnight))

        start_time = midnight + minutes_from_midnight * 60

        sec_meeting_card = meeting_card.find_element_by_css_selector("div")
        meeting_name = sec_meeting_card.get_attribute("title").replace("\n", " ")

        meeting_id = sec_meeting_card.get_attribute("id")

        meetings.append(Meeting(meeting_id, start_time, meeting_name, calendar_meeting=True))


def decide_meeting():
    newest_meetings = []

    meetings.sort( key=lambda x: x.time_started, reverse=True )
    newest_time = meetings[0].time_started

    for meeting in meetings:
        if meeting.time_started >= newest_time:
            newest_meetings.append( meeting )
        else:
            break

    if (current_meeting is None or newest_meetings[0].time_started > current_meeting.time_started) and (
            current_meeting is None or newest_meetings[0].m_id != current_meeting.m_id) and newest_meetings[
        0].m_id not in already_joined_ids:
        return newest_meetings[0]

    return


def join_meeting(meeting):
    global hangup_thread, current_meeting, already_joined_ids, active_correlation_id

    hangup()

    if meeting.calendar_meeting:
        switch_to_calendar_tab()
        join_btn = wait_until_found( f"div[id='{meeting.m_id}'] > div > button", 5 )

    else:
        browser.execute_script( f'window.location = "{conversation_link}a?threadId={meeting.channel_id}&ctx=channel";' )
        switch_to_teams_tab()

        join_btn = wait_until_found( f"div[id='{meeting.m_id}'] > calling-join-button > button", 5 )

    if join_btn is None:
        return

    browser.execute_script( "arguments[0].click()", join_btn )

    join_now_btn = wait_until_found( "button[data-tid='prejoin-join-button']", 30 )
    if join_now_btn is None:
        return

    uuid = re.search( uuid_regex, join_now_btn.get_attribute( "track-data" ) )
    if uuid is not None:
        active_correlation_id = uuid.group( 0 )
    else:
        active_correlation_id = ""

    # turn camera off
    video_btn = browser.find_element_by_css_selector( "toggle-button[data-tid='toggle-video']>div>button" )
    video_is_on = video_btn.get_attribute( "aria-pressed" )
    if video_is_on == "true":
        video_btn.click()
        print( "Video disabled" )

    # turn mic off
    audio_btn = browser.find_element_by_css_selector( "toggle-button[data-tid='toggle-mute']>div>button" )
    audio_is_on = audio_btn.get_attribute( "aria-pressed" )
    if audio_is_on == "true":
        audio_btn.click()
        print( "Microphone off" )

    if 'random_delay' in config and config['random_delay']:
        delay = random.randrange( 10, 31, 1 )
        print( f"Wating for {delay}s" )
        time.sleep( delay )

    # if 'random_delay' in config:
    #     if isinstance( config['random_delay'], bool ):
    #         print( f"Please update the random_delay in config.json file as per latest instructions in README" )
    #         if config['random_delay']:
    #             delay = random.randrange( 10, 31, 1 )
    #         else:
    #             delay = 0
    #     else:
    #         delay = random.randrange( config['random_delay'][0], config['random_delay'][1] + 1, 1 )
    #
    #     if delay > 0:
    #         print( f"Wating for {delay}s" )
    #         time.sleep( delay )

    # find again to avoid stale element exception
    join_now_btn = wait_until_found( "button[data-tid='prejoin-join-button']", 5 )
    if join_now_btn is None:
        return
    join_now_btn.click()

    current_meeting = meeting
    already_joined_ids.append( meeting.m_id )

    # if "join_message" in config and config["join_message"] != "":
    #     time.sleep( 3 )
    #     try:
    #         browser.execute_script( "document.getElementById('chat-button').click()" )
    #         text_input = wait_until_found( 'div[role="textbox"] > div', 5 )
    #
    #         js_change_text = """
    #           var elm = arguments[0], txt = arguments[1];
    #           elm.innerHTML = txt;
    #           """
    #
    #         browser.execute_script( js_change_text, text_input, config["join_message"] )
    #
    #         time.sleep( 3 )
    #         send_button = wait_until_found( "#send-message-button", 5 )
    #         send_button.click()
    #         print( f'Sent message {config["join_message"]}' )
    #         discord_notification( "Sent message", {config["join_message"]} )
    #     except (exceptions.JavascriptException, exceptions.ElementNotInteractableException):
    #         print( "Failed to send join message" )
    #         pass
    #
    print( f"Joined meeting: {meeting.title}" )
    # discord_notification( "Joined meeting", f"{meeting.title}" )

    if 'auto_leave_after_min' in config and config['auto_leave_after_min'] > 0:
        hangup_thread = Timer( config['auto_leave_after_min'] * 60, hangup )
        hangup_thread.start()


# def join_meeting(meeting):
#     global hangup_thread, current_meeting, already_joined_ids, active_correlation_id,joined
#
#     joined=0
#         # try:
#         #     check_element = browser.find_element_by_class_name("ts-calling-unified-bar")
#         # finally:
#         #     pass
#         # if check_element is not None:
#         #     print("no one in class")
#     hangup()
#
#     if meeting.calendar_meeting:
#         switch_to_calendar_tab()
#         join_btn = wait_until_found(f"div[id='{meeting.m_id}'] > div > button", 5)
#         if join_btn is None:
#             return
#
#         browser.execute_script("arguments[0].click()", join_btn)
#         #join_btn.click()
#
#     else:
#         browser.execute_script(f"window.location = 'https://teams.microsoft.com/_#/pre-join-calling/{meeting.m_id}';")
#
#     join_now_btn = wait_until_found("button[data-tid='prejoin-join-button']", 30)
#     if join_now_btn is None:
#         return
#
#     uuid = re.search(uuid_regex, join_now_btn.get_attribute("track-data"))
#     if uuid is not None:
#         active_correlation_id = uuid.group(0)
#     else:
#         active_correlation_id = ""
#
#     # turn camera off
#     video_btn = browser.find_element_by_css_selector("toggle-button[data-tid='toggle-video']>div>button")
#     video_is_on = video_btn.get_attribute("aria-pressed")
#     if video_is_on == "true":
#         video_btn.click()
#
#     # turn mic off
#     audio_btn = browser.find_element_by_css_selector("toggle-button[data-tid='toggle-mute']>div>button")
#     audio_is_on = audio_btn.get_attribute("aria-pressed")
#     if audio_is_on == "true":
#         audio_btn.click()
#         joined=1
#
#     if 'random_delay' in config and config['random_delay']:
#         delay = random.randrange(10, 31, 1)
#         print(f"Wating for {delay}s")
#         time.sleep(delay)
#
#     # find again to avoid stale element exception
#     join_now_btn = wait_until_found("button[data-tid='prejoin-join-button']", 5)
#     if join_now_btn is None:
#         return
#     join_now_btn.click()
#
#     current_meeting = meeting
#     already_joined_ids.append(meeting.m_id)
#
#     print(f"Joined meeting: {meeting.title}")
#     joined=1
#
#     if mode != 3:
#         switch_to_teams_tab()
#     else:
#         switch_to_calendar_tab()
#
#     if 'auto_leave_after_min' in config and config['auto_leave_after_min'] > 0:
#         hangup_thread = Timer(config['auto_leave_after_min'] * 60, hangup)
#         hangup_thread.start()
#
# def get_meeting_members():
#     meeting_elems = browser.find_elements_by_css_selector(".one-call")
#     for meeting_elem in meeting_elems:
#         try:
#             meeting_elem.click()
#             break
#         except:
#             continue
#
#     time.sleep(2)
#     browser.execute_script("document.getElementById('roster-button').click()")
#     wait_until_found(".ts-meeting-panel-components:not(.hide-meetings-panel)", 5, print_error=False)
#
#     participants = browser.find_elements_by_css_selector("calling-roster-section[ng-show*='participantsInCall'] li.vs-repeat-repeated-element")
#
#     if mode != 3:
#         switch_to_teams_tab()
#     else:
#         switch_to_calendar_tab()
#
#     return len(participants)

def hangup():
    global current_meeting, active_correlation_id
    if current_meeting is None:
        return

    try:
        hangup_btn = browser.find_element_by_css_selector( "button[data-tid='call-hangup']" )
        hangup_btn.click()

        print( f"Left Meeting: {current_meeting.title}" )

        current_meeting = None

        if hangup_thread:
            hangup_thread.cancel()

        return True
    except exceptions.NoSuchElementException:
        return False


def restart():
    timestamp = datetime.now()
    print( f"Restarting current script at [{timestamp:%H:%M:%S}]" )
    browser.close()
    browser.quit()

    meetings.pop( -1 )
    time.sleep( 5 )
    load_config()
    main()


def check(hangup):
    tim = []
    h = hangup
    pattern = '[0-9]{1,2}:[0-9]{2,2}'
    p2 = '[0-9]{1,2}'

    if len( meetings ) > 0:
        try:
            p = re.findall( pattern, current_meeting.title )
            print( p )
            timestamp = datetime.now()
            h = timestamp.hour
            m = timestamp.minute
            t = time.strptime( str( h ), "%H" )
            timevalue_12hour = str( time.strftime( "%I", t ) )
            h = int( timevalue_12hour )
            clhr = re.findall( p2, str( p ) )
            print( "current time", h, ':', m, "\n class hours : ", clhr )
        except:
            print( "fkkk ttt" )
            pass
        try:
            check_element = browser.find_element_by_class_name( "ts-calling-unified-bar" )
            if check_element is not None:
                print( check_element )
                print( "Class is joined" )
            elif h == 1:
                if int( h ) == int( clhr[0] ):
                    if int( m + 5 ) > int( clhr[1] ):
                        meetings.clear()
                        clhr.pop( -1 )
                        clhr.clear()
                        print( "RESTARTING" )

                        restart()
                    elif int( m > int( clhr[3] ) ):
                        print( "class hour has ended" )
                    else:
                        print( "not the time to join" )
        except exceptions.NoSuchElementException:
            print( "element not found" )
            print( "class not joined" )
            restart()
            if int( h ) == int( clhr[0] ) and int( m + 5 ) > int( clhr[1] ):
                restart()
            pass

    # for meeting in meetings:
    #
    #     print(meeting.time_started,type(meeting.time_started))
    #     if int(timestamp) >  meeting.time_started:
    #         timelapsed = 1
    #
    # if timelapsed ==1:
    #     print("class not joined")
    #     restart()


def isConnected():
    try:
        # reachable
        sock = socket.create_connection( ("1.1.1.1", 80) )
        if sock is not None:
            print( sock.type )
            sock.close()

        return True
    except OSError:
        pass
    return False


def death():
    recon = False
    while 1:
        timestamp = datetime.now()
        ch = 0
        window = gw.getWindowsWithTitle( 'Microsoft Teams - Google Chrome' )
        ch = len( window )
        if isConnected() == False:
            print( "Lost Internet" )
            time.sleep( 20 )

        print( "windows running : ", ch, window )
        try:
            if recon is True and isConnected() == True:
                print( "Quiting Current browser " )
                restart()
                recon = False
            else:
                if ch <= 20 and isConnected() == True:
                    print( f"Startingscript at [{timestamp:%H:%M:%S}]" )
                    main()


        except requests.exceptions.ConnectionError:
            print( "Something's nasty happened at ", time.localtime().tm_hour, ':', time.localtime().tm_min,
                   "waiting for shit to fix it self" )
            if isConnected() == False:
                recon = True
                print( "Lost Internet at ", time.localtime().tm_hour, ':', time.localtime().tm_min, " Rejoin" )
                time.sleep( 5 )
                pass


        except selenium.common.exceptions.WebDriverException:

            print( "Something's nasty happened at ", time.localtime().tm_hour, ':', time.localtime().tm_min,
                   "waiting for shit to fix it self" )
            if isConnected() == False:
                recon = True
                print( "Lost Internet at ", time.localtime().tm_hour, ':', time.localtime().tm_min, " Rejoin" )
                time.sleep( 5 )
                pass


def getmeet():
    global mt
    if len( meetings ) >= 0:
        print( "looking for meetings: " )

        for meeting in meetings:
            print( meeting, len( meetings ), "time :", meeting.time_started )
            mt = meeting.time_started
            print( "time :", type( mt ), mt )

    else:
        print( "no current meeting" )


def get_meeting_members():
    global current_meeting

    meeting_elems = browser.find_elements_by_css_selector( ".one-call" )

    # meeting has been closed by host
    if len( meeting_elems ) == 0:
        current_meeting = None
        print( "You are no longer in any meeting" )
        return

    for meeting_elem in meeting_elems:
        try:
            meeting_elem.click()
            break
        except:
            continue

    time.sleep( 2 )

    # open the meeting member side page
    try:
        browser.execute_script( "document.getElementById('roster-button').click()" )
    except exceptions.JavascriptException:
        print( "Failed to open meeting member page" )
        return None

    participants_elem = wait_until_found( "calling-roster-section[section-key='participantsInCall'] .roster-list-title",
                                          2, print_error=False )
    attendees_elem = wait_until_found( "calling-roster-section[section-key='attendeesInMeeting'] .roster-list-title", 2,
                                       print_error=False )

    if participants_elem is None and attendees_elem is None:
        print( "Failed to get meeting members" )
        return None

    if participants_elem is not None:
        participants = [int( s ) for s in participants_elem.get_attribute( "aria-label" ).split() if s.isdigit()]
    else:
        participants = [0]

    if attendees_elem is not None:
        attendees = [int( s ) for s in attendees_elem.get_attribute( "aria-label" ).split() if s.isdigit()]
    else:
        attendees = [0]

    # close the meeting member side page, this only makes a difference if pause_search is true
    try:
        browser.execute_script( "document.getElementById('roster-button').click()" )
    except exceptions.JavascriptException:
        # if the roster button doesn't exist click the three dots button before
        try:
            browser.execute_script( "document.getElementById('callingButtons-showMoreBtn').click()" )
            time.sleep( 1 )
            browser.execute_script( "document.getElementById('roster-button').click()" )
        except exceptions.JavascriptException:
            print( "Failed to close meeting member page, this might result in an error on next search" )

    return sum( participants + attendees )


def main():
    global config, meetings, mode, conversation_link, joined
    joined = r = 0

    mode = 1

    if isConnected() == True:
        if "meeting_mode" in config and 0 < config["meeting_mode"] < 4:
            mode = config["meeting_mode"]

        init_browser()

        browser.get( "https://teams.microsoft.com" )

        if config['email'] != "" and config['password'] != "":
            login_email = wait_until_found( "input[type='email']", 30 )
            if login_email is not None:
                login_email.send_keys( config['email'] )

            # find the element again to avoid StaleElementReferenceException
            login_email = wait_until_found( "input[type='email']", 5 )
            if login_email is not None:
                login_email.send_keys( Keys.ENTER )

            login_pwd = wait_until_found( "input[type='password']", 10 )
            if login_pwd is not None:
                login_pwd.send_keys( config['password'] )

            # find the element again to avoid StaleElementReferenceException
            login_pwd = wait_until_found( "input[type='password']", 5 )
            if login_pwd is not None:
                login_pwd.send_keys( Keys.ENTER )

            keep_logged_in = wait_until_found( "input[id='idBtn_Back']", 5 )
            if keep_logged_in is not None:
                keep_logged_in.click()

            use_web_instead = wait_until_found( ".use-app-lnk", 10, print_error=False )
            if use_web_instead is not None:
                use_web_instead.click()

        # if additional organisations are setup in the config file
        if 'organisation_num' in config and config['organisation_num'] > 1:
            change_organisation( config['organisation_num'] )

        print( "Waiting for correct page...", end='' )

        if wait_until_found( "#teams-app-bar", 5 ) is None:
            try:
                window = gw.getWindowsWithTitle( 'Microsoft Teams - Google Chrome' )
                print( window[0] )


            except:
                print( "Yo this fuck up again" )
                pyautogui.moveTo( 279, 264 )
                pyautogui.PAUSE = 3
                pyautogui.doubleClick()
                pyautogui.PAUSE = 1.5
                pyautogui.press( 'f5' )
                pyautogui.PAUSE = 1.5
                pyautogui.press( 'f5' )
                pyautogui.PAUSE = 1.5
                pyautogui.press( 'f5' )

        print( "\rFound page, do not click anything on the webpage from now on." )
        # wait a bit so the meetings are initialized
        time.sleep( 5 )

        if mode != 2:
            prepare_page( include_calendar=True )
        else:
            prepare_page( include_calendar=False )

        if mode != 3:
            switch_to_teams_tab()

            url = browser.current_url
            url = url[:url.find( "?" )]
            conversation_link = url

            teams = get_all_teams()

            if len( teams ) == 0:
                print(
                    "Not Teams found, is MS Teams in list mode? (switch to mode 3 if you only want calendar meetings)" )
                exit( 1 )

            for team in teams:
                print( team )

        check_interval = 10
        if "check_interval" in config and config['check_interval'] > 1:
            check_interval = config['check_interval']

        interval_count = 0
        switch_to_calendar_tab()

        # loop###
        while 1:
            pattern = '[0-9]{1,2}:[0-9]{2,2}'
            p2 = '[0-9]{1,2}'

            if isConnected() == True:
                timestamp = datetime.now()
                print( f"\n[{timestamp:%H:%M:%S}] Looking for new meetings" )

                if mode != 3:
                    switch_to_teams_tab()
                    teams = get_all_teams()

                    if len( teams ) == 0:
                        print( "Nothing found, is Teams in list mode?" )
                        exit( 1 )
                    else:
                        get_meetings( teams )

                if mode != 2:
                    print( mode )
                    switch_to_calendar_tab()
                    get_calendar_meetings()

                if len( meetings ) > 0:

                    print( "Found meetings: " )

                    for meeting in meetings:
                        print( meeting, len( meetings ), "time :", meeting.time_started )

                    cm = meetings[0]
                    # meeting_to_join = meetings[0]
                    meeting_to_join = decide_meeting()
                    if meeting_to_join is not None:
                        print( "attempting" )
                        joined = 0
                        try:
                            p = re.findall( pattern, cm.title )
                            print( p )
                            timestamp = datetime.now()
                            h = timestamp.hour
                            m = timestamp.minute
                            t = time.strptime( str( h ), "%H" )
                            timevalue_12hour = str( time.strftime( "%I", t ) )
                            h = int( timevalue_12hour )
                            clhr = re.findall( p2, str( p ) )
                            print( "current time", h, ':', m, "\n class hours : ", clhr )
                            if h is not None:
                                if int( h ) <= int( clhr[2] ):
                                    if int( m ) <= int( clhr[1] ) + 5:

                                        join_meeting( meeting_to_join )
                                    elif int( h ) > int( clhr[2] ) and int( m ) <= int( clhr[3] ) + 5:
                                        print( "Class time has elapsed" )

                                else:
                                    print( "Waiting for right Time . " )
                        except exceptions:
                            print( "Waiting to recive meeting time" )

                    else:
                        print( "Meeting : ", meetings[0], " is joined" )
                        joined = 1

                        #
                        # members = get_meeting_members()
                        # if current_meeting is None:
                        #     continue
                        # print( "no of members present : ", members )
                        #
                        # if members is not None and members <= 5:
                        #     print( f"\n Too few members in meeting ,Exiting @ [{timestamp:%H:%M:%S}]" )
                        #     hangup()
                else:
                    joined = 0
                    print( "no meeting found" )
                    r = r + 5
                    print( "Restart @ 2000 :", r )
                    if r > 2000:
                        r = 0
                        restart()

                        # check(hangup=0)

                meetings = []

                if "leave_if_last" in config and config[
                    'leave_if_last'] and interval_count % 5 == 0 and interval_count > 0:
                    if current_meeting is not None:
                        members = get_meeting_members()

                        if members <= 1:
                            print( members )

                            hangup()
                            interval_count = 0
                            check( hangup=1 )

                interval_count += 1

                time.sleep( check_interval )

                # try:
                #     if meeting_to_join is not None and browser is not None :
                #
                #         if meeting.time_started is not None and joined ==1:
                #             if members is None:
                #                 print( f"\n Bot is not working fine , Restarting@ [{timestamp:%H:%M:%S}]" )
                #                 browser.quit()
                #                 browser.close()
                #
                #         print("time ",meeting.time_started," type : ",type(meeting.time_started))
                # except exceptions:
                #     print(exceptions)


            else:
                recconnect = 1
                print( "Waiting for connection" )
                time.sleep( 5 )
                if browser is not None:
                    print( "quiting browser" )

    elif isConnected() != True:
        print( "No internet connection found" )


if __name__ == "__main__":

    getmeet()
    timestamp = datetime.now()
    print( f"\nStarting Script at : [{timestamp:%H:%M:%S}] " )
    try:
        if 'run_at_time' in config and config['run_at_time'] != " ":
            now = datetime.now()
            run_at = datetime.strptime( config['run_at_time'], "%H:%M" ).replace( year=now.year, month=now.month,
                                                                                  day=now.day )

            if run_at.time() < now.time():
                run_at = datetime.strptime( config['run_at_time'], "%H:%M" ).replace( year=now.year,
                                                                                      month=now.month,
                                                                                      day=now.day + 1 )

            start_delay = (run_at - now).total_seconds()

            print( f"Waiting until {run_at} ({int( start_delay )}s)" )
            time.sleep( start_delay )
    except:
        print( "Run at time not found in config" )

    start = 1
    recconnect = 0
    load_config()
    # print(meetings)
    win = th.ThemedTk()
    win.get_themes()
    win.set_theme( "clearlooks" )
    win.title( "Auto Attendance" )
    win.geometry( "900x380" )
    ar = Label( win, text="Enter UUID :" )
    sg = Label( win, text="Enter Password :" )
    ui = Entry( win )
    passw = Entry( win )
    ar.grid( row=4, column=0 )
    sg.grid( row=5, column=0 )
    ui.grid( row=4, column=4 )
    passw.grid( row=5, column=4 )
    br = Label( win, text="Developed by B#", fg="red", height=3 )
    br.grid( row=7, column=4, columnspan=1 )
    cr = Label( win, text="Auto join meetings and Classes", fg="black", bg="green" )
    cr.grid( row=10, column=4, columnspan=1 )
    btn = Button( win, text="Run", bg="green", fg="white", command=main, height=1, width=20 )
    btn.grid( row=8, column=4 )
    image = PhotoImage( file="p2.png" )
    bl = Label( win, image=image )
    bl.grid( row=11, column=4, columnspan=1 )
    bl.config( height=200, width=750 )
    print( recconnect )
    print( "affter mainloop:", recconnect )
    win.mainloop()

    # def fails():
    #
    #         if isConnected() == True:
    #
    #             global mt
    #             getmeet()
    #             timestamp = datetime.now()
    #             start = main()
    #
    #             print( f"\nStarting Script at : [{timestamp:%H:%M:%S}] " )
    #
    #             if 'run_at_time' in config and config['run_at_time'] != "":
    #                 now = datetime.now()
    #                 run_at = datetime.strptime( config['run_at_time'], "%H:%M" ).replace( year=now.year, month=now.month,
    #                                                                                       day=now.day )
    #
    #                 if run_at.time() < now.time():
    #                     run_at = datetime.strptime( config['run_at_time'], "%H:%M" ).replace( year=now.year,
    #                                                                                           month=now.month,
    #                                                                                           day=now.day + 1 )
    #
    #                 start_delay = (run_at - now).total_seconds()
    #
    #                 print( f"Waiting until {run_at} ({int( start_delay )}s)" )
    #                 time.sleep( start_delay )
    #
    #
    #
    #                 # try:
    #                 #     death()
    #                 # except exceptions.TimeoutException:
    #                 #     print( "x R.I.P x" )
    #                 # finally:
    #                 #     print( exceptions )

    # elif recconnect == 1:
    #             print( f"\n[{timestamp:%H:%M:%S}] Restarting due to internet loss" )
    #             time.sleep( 5 )
    #             load_config()
    #             main()
    #             recconnect = 0

    # if isConnected() == True and recconnect == 1:
    #     #invoke
    #     recconnect =0
    # elif isConnected()!=True:
    #     recconnect = 1
    #     print( "waiting for connection" )
    #     print("on loop:",recconnect)
    #     time.sleep(5)
    #     win.quit()
    #
    #     time.sleep(5)
    #
    # else:
    #     print("bot is alright so far ")
    #     time.sleep( 5 )
    #     recconnect = 0
# # while 1:
# #
# #
# #
# #         global mt
# #
# #         recconnect=0
# #         getmeet()
# #         timestamp = datetime.now()
# #
# #
# #
# #         if isConnected() == True and recconnect == 0:
# #             print(f"\nStarting Script at : [{timestamp:%H:%M:%S}] ")
# #             start = main()
# #             if 'run_at_time' in config and config['run_at_time'] != "":
# #                 now = datetime.now()
# #                 run_at = datetime.strptime( config['run_at_time'], "%H:%M" ).replace( year=now.year, month=now.month,
# #                                                                                       day=now.day )
# #
# #                 if run_at.time() < now.time():
# #                     run_at = datetime.strptime( config['run_at_time'], "%H:%M" ).replace( year=now.year,
# #                                                                                           month=now.month,
# #                                                                                           day=now.day + 1 )
# #
# #                 start_delay = (run_at - now).total_seconds()
# #
# #                 print( f"Waiting until {run_at} ({int( start_delay )}s)" )
# #                 time.sleep( start_delay )
# #
# #             try:
# #                 death()
# #             except exceptions.TimeoutException:
# #                 print( "x R.I.P x" )
# #             finally:
# #                 print( exceptions )
# #
# #             if recconnect == 1:
# #                 print(f"\n[{timestamp:%H:%M:%S}] Restarting due to internet loss")
# #                 time.sleep(5)
# #                 load_config()
# #                 main()
# #                 recconnect = 0
# #
# #
# #
# #         else:
# #             recconnect = 1
# #             print("Internet lost")
# #             time.sleep(10)
#
# # try:
# #     window = gw.getWindowsWithTitle('Untitled')[0]
# #     pyautogui.press('f5')
# #     break
# # except:
# #     continue
