import pandas as pd
from playwright.sync_api import sync_playwright
import time
import urllib.parse

def process_codes(codes_list):
    """
    Takes a list of product codes, scrapes Kohler.com, and returns a DataFrame with results.
    """
    # Create DataFrame from list
    df = pd.DataFrame({'Code': codes_list})
    
    # Add columns
    df['Product Name'] = ''
    df['Link'] = ''
    df['Match Verified'] = False

    with sync_playwright() as p:
        # Launch browser (headless=True for reliable execution in this environment)
        browser = p.chromium.launch(headless=True)
        # Use a real user agent to avoid immediate blocking
        context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        page = context.new_page()

        for index, row in df.iterrows():
            code = str(row['Code']).strip()
            if not code:
                continue
            
            print(f"Processing: {code}")
            
            # Parse code to get base and color
            # Assumption: Color is always after the last hyphen
            if '-' in code:
                base_code = code.rsplit('-', 1)[0]
                color_code = code.rsplit('-', 1)[1]
            else:
                base_code = code
                color_code = None
                
            print(f"Base Code: {base_code}, Color Code: {color_code}")

            try:
                # Direct Search on Kohler.com
                print(f"Navigating to Kohler homepage...")
                page.goto("https://www.kohler.com/en")
                page.wait_for_load_state('domcontentloaded')

                # Click search icon
                try:
                    page.click("button[aria-label='Search']")
                    # Type code into search input
                    page.fill("input#search-side-panel__search-control", code)
                    # Press Enter
                    page.keyboard.press("Enter")
                    
                    # Wait for navigation (URL change) or results
                    page.wait_for_url("**/products/**", timeout=15000)
                    print(f"Navigated to product page: {page.url}")
                    
                except Exception as e:
                    print(f"Search failed or timed out: {e}")
                    df.at[index, 'Product Name'] = "Search Failed"
                    df.at[index, 'Link'] = "Search Failed"
                    df.at[index, 'Match Verified'] = False
                    continue

                # Handle Color Selection
                current_url = page.url
                if color_code and color_code not in current_url:
                     print(f"URL does not contain color code {color_code}. Attempting to select color...")
                     try:
                        color_code_upper = color_code.upper()
                        # Selector for the input element
                        color_input_selector = f'input[id$="-{color_code_upper}"], input[value$="-{color_code_upper}"]'
                        
                        # Check if such an input exists
                        if page.locator(color_input_selector).count() > 0:
                            color_input = page.locator(color_input_selector).first
                            color_swatch = color_input.locator('xpath=./ancestor::div[@role="radio"]')
                            
                            if color_swatch.count() > 0:
                                print(f"Found swatch for {color_code_upper}. Clicking...")
                                color_swatch.first.click()
                                try:
                                    page.wait_for_url(f"**/*{color_code_upper}*", timeout=5000)
                                    print("URL updated.")
                                except:
                                    print("URL did not update or timed out.")
                            else:
                                print(f"Found input for {color_code_upper} but could not find clickable parent swatch.")
                        else:
                            print(f"Color option {color_code_upper} not found on page.")
                     except Exception as e:
                        print(f"Error selecting color: {e}")

                # Extract Product Name (H1)
                try:
                    product_name = page.locator('h1').first.inner_text(timeout=5000)
                except:
                    product_name = "Name not found"
                
                # Verify if the displayed product code matches the requested code
                is_match = False
                try:
                    # Normalize spaces
                    page_text = page.locator('body').inner_text()
                    if code in page_text:
                        is_match = True
                        print(f"Verification Successful: Found {code} on page.")
                    else:
                        print(f"Verification Warning: Could not find exact code {code} on page.")
                except:
                    pass

                print(f"Product Name: {product_name}")
                
                # Update DataFrame
                df.at[index, 'Product Name'] = product_name
                df.at[index, 'Link'] = page.url # Save the actual current URL which should include the color
                df.at[index, 'Match Verified'] = is_match

            except Exception as e:
                print(f"Error processing {code}: {e}")
                df.at[index, 'Product Name'] = "Error"
                df.at[index, 'Link'] = str(e)
                df.at[index, 'Match Verified'] = False
            
            # Be nice to Google
            time.sleep(2)

        browser.close()
    
    return df

def scrape_kohler():
    input_file = 'input.xlsx'
    output_file = 'output.xlsx'
    
    try:
        input_df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"Error: {input_file} not found.")
        return

    if 'Code' not in input_df.columns:
        print("Error: 'Code' column not found in input file.")
        return

    codes = input_df['Code'].tolist()
    result_df = process_codes(codes)
    
    # Merge results back if needed, or just save the result_df
    # For simplicity, let's just save the result_df which has the codes and results
    result_df.to_excel(output_file, index=False)
    print(f"Done. Results saved to {output_file}")

if __name__ == "__main__":
    scrape_kohler()
