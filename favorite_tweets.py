import tweepy as tw

auth = tw.OAuthHandler('consumer_key','consumer_secret') #insert consumer_key and consumer_secret
auth.set_access_token('access_token','access_token_secret') #insert access_token and access_token_secret


api = tw.API(auth, wait_on_rate_limit=True)

input_user = input("Enter Your Twitter Username: ")


user_favs=api.favorites(input_user,count=200)

for t in user_favs:
	print(t.text)