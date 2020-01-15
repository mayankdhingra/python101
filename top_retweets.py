import tweepy as tw
import operator

auth = tw.OAuthHandler('consumer_key','consumer_secret') #insert consumer_key and consumer_secret
auth.set_access_token('access_token','access_token_secret') #insert access_token and access_token_secret

#testcomment
api = tw.API(auth, wait_on_rate_limit=True)

input_user = input("Enter Your Twitter Username Only: ")


#user_favs=api.favorites(input_user,count=200)
user_rts=api.retweets_of_me(count=200)
top_rts = {}

for t in user_rts:
	if t.retweet_count>0:
		top_rts.update({t.text:t.retweet_count})
top_rts=dict(sorted(top_rts.items(), key=operator.itemgetter(1), reverse=True))
print(top_rts)
