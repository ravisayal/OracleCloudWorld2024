#!/bin/bash

# Source the parameter file with sensitive values
source ./params.conf

# URL and headers for the request
url='https://events.rf.oracle.com/api/search'
headers=(
  '-H' 'Accept: */*'
  '-H' 'Accept-Language: en-US,en;q=0.9'
  '-H' 'Connection: keep-alive'
  '-H' 'Content-Type: application/x-www-form-urlencoded; charset=UTF-8'
  '-H' 'Origin: https://reg.rf.oracle.com'
  '-H' 'Referer: https://reg.rf.oracle.com/'
  '-H' 'Sec-Fetch-Dest: empty'
  '-H' 'Sec-Fetch-Mode: cors'
  '-H' 'Sec-Fetch-Site: same-site'
  '-H' 'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36'
  '-H' "rfApiProfileId: ${rfApiProfileId}"
  '-H' "rfAuthToken: ${rfAuthToken}"
  '-H' "rfWidgetId: ${rfWidgetId}"
  '-H' 'sec-ch-ua: "Not)A;Brand";v="99", "Google Chrome";v="127", "Chromium";v="127"'
  '-H' 'sec-ch-ua-mobile: ?0'
  '-H' 'sec-ch-ua-platform: "Windows"'
)

# Start the loop from 0 to 1042 in increments of 50
for (( i=0; i<=1042; i+=50 )); do
  # Print the current iteration
  echo "Fetching records starting from $i"

  # Make the curl request and save the output to a file, with the --insecure option to bypass SSL issues
  curl --insecure "${url}" \
    "${headers[@]}" \
    --data-raw "search=&type=session&browserTimezone=America%2FNew_York&catalogDisplay=list&from=${i}" \
    -o "output_${i}.json"

  # Pause briefly between requests to avoid overloading the server
  sleep 1
done

echo "Data fetching completed."
