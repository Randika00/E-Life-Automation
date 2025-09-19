import re
import os
from datetime import datetime
import requests
import common_function
import subprocess
import pdfkit
from pathlib import Path
import pandas as pd
from bs4 import BeautifulSoup

WKHTMLTOPDF_URL = "https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox-0.12.6-1.mxe-cross-win64.7z"
SEVEN_ZIP_PATH = "C:\\Program Files\\7-Zip\\7z.exe"  # Path to 7z.exe
EXTRACTION_DIR = os.path.join(os.getcwd(), "wkhtmltopdf_bin")  # Persistent extraction folder
ARCHIVE_PATH = os.path.join(EXTRACTION_DIR, "wkhtmltox.7z")

def get_soup(url):
    global statusCode
    response = requests.get(url,headers=headers,stream=True)
    statusCode = response.status_code
    soup= BeautifulSoup(response.content, 'html.parser')
    return soup

def get_ordinal_suffix(n):
    if 11 <= n % 100 <= 13:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return str(n) + suffix

def download_pdf(article_details,pdf_file_path):
    try:
        pre_pdf_link = article_details.find("a", class_="button button--icon-download button--action")
        if pre_pdf_link:
            pdf_link = pre_pdf_link["href"]
        else:
            pdf_link = article_details.find("a", string="Article PDF")["href"]
    except Exception:
        raise Exception("PDF link could not be found")

    pdf_response = requests.get(pdf_link,headers=headers)

    if pdf_response.status_code == 200:

        print("⏳ Wait until the PDF is downloaded")
        pdf_content = pdf_response.content
        with open(pdf_file_path, "wb") as file:
            file.write(pdf_content)
        print(f"📁 PDF saved at: {pdf_file_path}"+"\n")

    else:
        raise Exception("Unable to download the PDF")

def setup_wkhtmltopdf():
    """Download and extract wkhtmltopdf only if not already available."""
    wkhtmltopdf_exe = None

    # Check if already extracted
    for root, dirs, files in os.walk(EXTRACTION_DIR):
        if "wkhtmltopdf.exe" in files:
            wkhtmltopdf_exe = os.path.join(root, "wkhtmltopdf.exe")
            break

    if wkhtmltopdf_exe and os.path.isfile(wkhtmltopdf_exe):
        print("wkhtmltopdf is already set up.")
        return wkhtmltopdf_exe

    print("Downloading wkhtmltopdf archive...")
    os.makedirs(EXTRACTION_DIR, exist_ok=True)
    with requests.get(WKHTMLTOPDF_URL, stream=True) as r:
        with open(ARCHIVE_PATH, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

    print("Extracting wkhtmltopdf with 7z.exe...")
    subprocess.run([
        SEVEN_ZIP_PATH,
        'x',
        ARCHIVE_PATH,
        f'-o{EXTRACTION_DIR}',
        '-y'
    ], check=True)

    # Find the binary again after extraction
    for root, dirs, files in os.walk(EXTRACTION_DIR):
        if "wkhtmltopdf.exe" in files:
            wkhtmltopdf_exe = os.path.join(root, "wkhtmltopdf.exe")
            break

    if not wkhtmltopdf_exe:
        raise FileNotFoundError("wkhtmltopdf.exe not found after extraction.")

    return wkhtmltopdf_exe


def convert_html_to_pdf(html_file_path, output_pdf_path,exe_path):
    """Convert a single HTML file to PDF using the prepared wkhtmltopdf binary."""
    try:
        config = pdfkit.configuration(wkhtmltopdf=exe_path)
        options = {
            'enable-local-file-access': True,
            'load-error-handling': 'ignore',
            'load-media-error-handling': 'ignore',
        }
        pdfkit.from_file(html_file_path, output_pdf_path, configuration=config,options=options)
    except Exception as e:
        pass

headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Priority": "u=0, i",
    "Sec-Ch-Ua": "\"Not/A)Brand\";v=\"8\", \"Chromium\";v=\"126\", \"Google Chrome\";v=\"126\"",
    "Sec-Ch-Ua-Arch": "\"x86\"",
    "Sec-Ch-Ua-Bitness": "\"64\"",
    "Sec-Ch-Ua-Full-Version": "\"126.0.6478.127\"",
    "Sec-Ch-Ua-Full-Version-List": "\"Not/A)Brand\";v=\"8.0.0.0\", \"Chromium\";v=\"126.0.6478.127\", \"Google Chrome\";v=\"126.0.6478.127\"",
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Model": "\"\"",
    "Sec-Ch-Ua-Platform": "\"Windows\"",
    "Sec-Ch-Ua-Platform-Version": "\"15.0.0\"",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
}

error_messages = {
    200: "Server error: Unable to find HTML content",
    400: "Error in the site: 400 Bad Request",
    401: "Error in the site: 401 Unauthorized",
    403: "Error in the site: Error 403 Forbidden",
    404: "Error in the site: 404 Page not found!",
    408: "Error in the site: Error 408 Request Timeout",
    500: "Error in the site: Error 500 Internal Server Error",
    526: "Error in the site: Error 526 Invalid SSL certificate"
}

statusCode = None
attachment= None
error_list = []
completed_list = []

def main():
    exe_path = setup_wkhtmltopdf()

    current_datetime = datetime.now()
    current_date = str(current_datetime.date())
    current_time = current_datetime.strftime("%H:%M:%S")

    ini_path = os.path.join(os.getcwd(), "Info.ini")
    Download_Path, Email_Sent, Check_duplicate, user_id = common_function.read_ini_file(ini_path)
    current_out = common_function.return_current_outfolder(Download_Path, user_id)
    out_excel_file = common_function.output_excel_name(current_out)

    url = "https://elifesciences.org/"
    base_url = "https://elifesciences.org"

    data = []
    try:
        try:
            archive_link = base_url + get_soup(url).find("a",string=lambda text:text and "Monthly archive" in text)["href"]
        except Exception:
            raise Exception("No archive link found")

        try:
            archive_soup = get_soup(archive_link)
            current_link = base_url + archive_soup.find("ol",class_="grid-listing").findAll("li",class_="grid-listing-item")[-1].find("a")["href"]
            if current_link:
                current_soup = get_soup(current_link)
            else:
                raise Exception("No link found for the current page")
        except Exception:
            raise Exception("No link found for the current page")

        try:
            all_articles = current_soup.find("h3", string="Research articles").find_next_sibling("ol").findAll("li",class_="listing-list__item")[:15]
            if all_articles:
                Total_count = len(all_articles)
                print(f"✅ Total number of articles:{Total_count}", "\n")
            else:
                raise Exception("No links found for articles")
        except Exception:
            raise Exception("No links found for articles")

        for index,sin_art in enumerate(all_articles):

            article_link = None
            article_number = index+1
            try:
                try:
                    article_link = base_url + sin_art.find("a",class_="teaser__header_text_link")["href"]
                except Exception:
                    raise Exception("No article link found")

                try:
                    article_title = sin_art.find("a",class_="teaser__header_text_link").get_text(strip=True)
                except Exception:
                    raise Exception("No article title found")

                try:
                    input_date = sin_art.find("time").get_text(strip=True)
                    parsed_date = datetime.strptime(input_date, "%b %d, %Y")
                    data_time = parsed_date.strftime("%d/%m/%Y")
                except Exception:
                    data_time = ""

                article_details = get_soup(article_link)
                if not statusCode ==200:
                    Error_message = error_messages.get(statusCode)

                    if Error_message is None:
                        Error_message = "Error in the article link"
                    raise Exception(Error_message)

                try:
                    doi_link = article_details.find("li",class_="descriptors__identifier")
                    if doi_link:
                        DOI = doi_link.find("a").get_text(strip=True).split("org/")[-1]
                    else:
                        DOI = article_details.find("a",class_="doi__link").get_text(strip=True).split("org/")[-1]
                except Exception:
                    raise Exception("DOI number could not be found")

                article_id = re.search(r"eLife.(\d+)", DOI).group(1)

                current_id_folder = os.path.join(current_out, f"{article_id}")

                if not os.path.exists(current_id_folder):
                    os.makedirs(current_id_folder)

                html_file_path = os.path.join(current_id_folder, f"{article_id}.html")
                pdf_file_path = os.path.join(current_id_folder, f"{article_id}.pdf")

                print("✅ " + get_ordinal_suffix(article_number) + " article details have been scraped")

                with open(html_file_path, 'w', encoding='utf-8') as file:
                    file.write(str(article_details))

                pdf_link = article_details.find("a",string="Download")
                if pdf_link:
                    download_pdf(article_details,pdf_file_path)
                else:
                    print("⏳ Wait until the PDF is downloaded")
                    convert_html_to_pdf(html_file_path, pdf_file_path,exe_path)
                    print(f"📁 PDF saved at: {pdf_file_path}" + "\n")

                data.append(
                    {"Title": article_title, "DOI": DOI, "URL": current_link,"Publication Date":data_time,
                     "SOURCE File PDF":f"{article_id}.pdf","SOURCE File HTML":f"{article_id}.html",
                     "Type of article":"PDF","User ID":user_id})

                df = pd.DataFrame(data)
                df.to_excel(out_excel_file, index=False)
                completed_list.append(f"{article_link}")

            except Exception as error:
                error_list.append(f"{article_link}")

        try:
            attachment_path = out_excel_file
            if os.path.isfile(attachment_path):
                attachment = attachment_path
            else:
                attachment = None
            common_function.attachment_for_email(error_list, completed_list,
                                                 len(completed_list), ini_path, attachment, current_date,
                                                 current_time)
        except Exception as error:
            message = f"Failed to send email : {str(error)}"
            common_function.email_body_html(current_date, current_time, error_list,
                                            completed_list,
                                            len(completed_list), attachment, current_out)

    except Exception as error:
        Error_message = error_messages.get(statusCode)
        if statusCode == 200 and str(error):
            Error_message = "Error in the site: " + str(error)

        if Error_message is None:
            Error_message = "Error in the site: " + str(error)

        error_list.append(Error_message)
        try:
            attachment_path = out_excel_file
            if os.path.isfile(attachment_path):
                attachment = attachment_path
            else:
                attachment = None
            common_function.attachment_for_email(error_list, completed_list,
                                                 len(completed_list), ini_path, attachment, current_date,
                                                 current_time)
        except Exception as error:
            common_function.email_body_html(current_date, current_time, error_list,
                                            completed_list,
                                            len(completed_list), attachment, current_out)

if __name__ == "__main__":
    main()