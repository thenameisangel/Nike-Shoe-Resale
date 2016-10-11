#!/usr/bin/env python
# encoding: utf-8

import tweepy, csv, random, datetime, time, xlsxwriter

#Twitter API credentials
ckey=""
csecret=""
atoken=""
asecret=""

auth = tweepy.OAuthHandler(ckey, csecret)
auth.set_access_token(atoken, asecret)
api = tweepy.API(auth)

def grab_tweets(keywords, start, end):

	all_tweets = []
	new_tweets = tweepy.Cursor(api.search, q = keywords, since = start, until = end).items()

	for obj in new_tweets:
		if not obj.retweeted and "RT @" not in obj.text:
			tweet = []
			tweet.append(obj.created_at.strftime('%m/%d/%Y @ %H:%M'))
			tweet.append(obj.user.screen_name)
			tweet.append(obj.text)
			tweet.append(obj.retweet_count)
			all_tweets.append(tweet)

	return(all_tweets)

def write_xls(all_tweets, search_terms, workbook):
	format01 = workbook.add_format()
	format02 = workbook.add_format()
	format03 = workbook.add_format()
	format01.set_align('center')
	format01.set_align('vcenter')
	format02.set_align('center')
	format02.set_align('vcenter')
	format03.set_align('center')
	format03.set_align('vcenter')
	format03.set_bold()

	header = ['Date', 'Username', 'Tweet', '# RT']

	title = "Keywords - "

	for elt in search_terms:
		title += elt + ' '

	worksheet = workbook.add_worksheet(title)

	out1 = all_tweets
	row = 0
	col = 0

	worksheet.set_column('A:A', 20)
	worksheet.set_column('B:B', 20)
	worksheet.set_column('C:C', 100)
	worksheet.set_column('D:D', 7)

	for item in header:
		worksheet.write(row, col, item, format03)
		col = col + 1

	row += 1
	col = 0

	for elt in out1:
		write = []
		write = [elt[0], elt[1], elt[2], elt[3]]

		format01.set_num_format('yyyy/mm/dd hh:mm:ss')
		worksheet.write(row, 0, write[0], format02)
		worksheet.write(row, 1, write[1], format02)
		worksheet.write(row, 2, write[2], format02)
		worksheet.write(row, 3, write[3], format02)
		row += 1
		col = 0

def main():

	#set search terms
	search_terms = "air jordan 10 nyc"
	start_date = "2016-04-25"
	end_date = "2016-04-27"

	#grab tweets
	tweet_lst = []
	tweet_lst = grab_tweets(search_terms, start_date, end_date)

	#write excel File
	workbook = xlsxwriter.Workbook('Twitter_data.xlsx')
	write_xls(tweet_lst, search_terms, workbook)
	workbook.close()

main()
