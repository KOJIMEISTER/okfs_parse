import requests
import time
import json
from typing import Optional

def post_request(data: dict, max_retries: int = 5) -> Optional[dict]:
    """
    Sends a POST request to the specified URL with the given data.
    Handles "Too Many Requests" by waiting and retrying.

    Args:
        data (dict): The JSON payload to send in the POST request.
        max_retries (int): Maximum number of retries for 429 errors.

    Returns:
        dict or None: The JSON response from the server if successful, else None.
    """
    headers = {
        "Host": "websbor.rosstat.gov.ru",
        "Origin": "https://websbor.rosstat.gov.ru",
        "Content-Type": "application/json",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Accept": "application/json, text/plain, */*",
        "Accept-language": "ru,en;q=0.9",
        "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/132.0.0.0 Mobile Safari/537.36"
    }

    url = "https://websbor.rosstat.gov.ru/webstat/api/gs/organizations"
    retries = 0

    while retries <= max_retries:
        try:
            response = requests.post(url, json=data, headers=headers, verify=False)
            response.raise_for_status()  # Raises HTTPError for bad responses

            # If the request is successful, return the JSON content
            return response.json()

        except requests.exceptions.HTTPError as http_err:
            if response.status_code == 429:
                retries += 1
                print(f"Too many requests. Retry {retries}/{max_retries} after 10 seconds.")
                time.sleep(10)
                continue  # Retry the request
            else:
                print(f"HTTP error occurred: {http_err} - Status Code: {response.status_code}")
                break  # Exit the loop for other HTTP errors

        except requests.exceptions.SSLError as ssl_err:
            print(f"SSL error occurred: {ssl_err}")
            break

        except requests.exceptions.RequestException as req_err:
            print(f"Request exception occurred: {req_err}")
            break

        except Exception as err:
            print(f"An unexpected error occurred: {err}")
            break

    print("Failed to retrieve the response after multiple retries.")
    return None

def extract_okfs_code(response_json: dict) -> Optional[str]:
    """
    Extracts the 'code' value from the 'okfs' field in the JSON response.

    Args:
        response_json (dict): The JSON response from the server.

    Returns:
        str or None: The 'code' value if found, else None.
    """
    try:
        # Assuming the response is a list of organizations
        for organization in response_json:
            okfs = organization.get("okfs", {})
            code = okfs.get("code")
            if code:
                return code
        print("'code' not found in any 'okfs' field.")
    except (TypeError, AttributeError) as parse_err:
        print(f"Error parsing JSON: {parse_err}")
    return None

if __name__ == "__main__":
    payload = {
        "inn": "614501324623"
    }

    response = post_request(payload)
    if response:
        # If the response is a list, process the first item or iterate as needed
        # Here, we'll assume it's a list and take the first organization's 'okfs' code
        if isinstance(response, list) and response:
            okfs_code = extract_okfs_code(response)
            if okfs_code:
                print(f"Extracted 'okfs' code: {okfs_code}")
            else:
                print("Could not extract 'okfs' code from the response.")
        else:
            print("Unexpected response format.")
    else:
        print("No response received from the server.")
