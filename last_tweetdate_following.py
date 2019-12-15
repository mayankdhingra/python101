import tweepy as tw
from datetime import datetime, timezone,timedelta

auth = tw.OAuthHandler('consumer_key','consumer_secret') #insert consumer_key and consumer_secret
auth.set_access_token('access_token','access_token_secret') #insert access_token and access_token_secret

api = tw.API(auth, wait_on_rate_limit=True)

input_user = input("Enter Your Twitter Username: ")

user = api.get_user(input_user)
try:
	user_friends=api.friends_ids(user)[:10]

	#print(user.screen_name)		
	print(f"{user.screen_name} has {user.followers_count} followers and {user.friends_count} friends")

	offset_hours = +5.5
#local_timestamp = clean_timestamp + timedelta(hours=offset_hours)
	for f in user_friends:
			friend=api.get_user(f)
		#print(f" following {friend.screen_name}, last tweet on {friend.status._json['created_at']}")
			print(f" following {friend.screen_name}, last tweet on {datetime.strptime(friend.status._json['created_at'], '%a %b %d %H:%M:%S +0000 %Y') + timedelta(hours=offset_hours) } ")
#tweets = tw.Cursor(api.search,q=search_word,lang="en",since=date_since).items(10)
#print(f"showing results for {search_word}")
#for tweet in tweets:
#	print(tweet.text)
except:
	print("For now this program only works for your username")

