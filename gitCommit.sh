git config --global user.email "jajohnsylvester@gmail.com"
git config --global user.name "John Sylvester"
git add .
git commit -m "Changed $(date +'%Y-%m-%d %H:%M:%S')"

# The Base64 encoded string
encoded_string="Z2hwXzJDZVc4YWdaMkNmRWdZN1dEUmZEbHpKVWNWQVVtcTJYQ0JLaA==="

# Decode the string and store it in a variable
decoded_variable=$(echo "$encoded_string" | base64 --decode)
git push -u https://jajohnsylvester:$decoded_variable@github.com/jajohnsylvester/jupiternotebook.git main

