#source ChatGPT 2023-11-14 > https://chat.openai.com/c/8b1270c7-f108-4462-bf43-0262ad7ebeec

import requests
import os
import time

def download_video(url, filename):
    try:
        existing_file_size = 0

        while True:
            headers = {}

            # Check if the file already exists
            if os.path.exists(filename):
                existing_file_size = os.path.getsize(filename)
                headers['Range'] = f'bytes={existing_file_size}-'

            with requests.get(url, stream=True, headers=headers) as response:
                response.raise_for_status()
                total_size_in_bytes = int(response.headers.get('content-length', 0)) + existing_file_size
                block_size = 1024  # 1 Kibibyte
                progress_bar_size = 50
                print(f"\nTotal file size: {total_size_in_bytes / 1024 / 1024:.2f} MB, Existing file size: {existing_file_size / 1024 / 1024:.2f} MB")

                with open(filename, 'ab') as file:  # Append to the file if it exists
                    for data in response.iter_content(block_size):
                        file.write(data)
                        progress_status = (file.tell() / total_size_in_bytes)
                        filled_length = int(progress_bar_size * progress_status)
                        bar = '█' * filled_length + '-' * (progress_bar_size - filled_length)
                        print(f"\rProgress: |{bar}| {progress_status * 100:.2f}%", end="")

                    # Check if download is complete within the 'with open' block
                    if file.tell() == total_size_in_bytes:
                        print("\nDownload complete.")
                        return f"Downloaded to: {filename}"

            # Wait before retrying
            print("\nIncomplete download, retrying...")
            time.sleep(5)

    except requests.exceptions.RequestException as e:
        return f"Error: {e}"

# URL of the video
video_url = "https://live.licdn.com/bitmovinneuprod/bitmovinneuprodoutputstoragecontainer/a0851f35-a001-444d-af2a-3448af3f49f5/RtmpLiveEncoding/video_1024.m3u8"

# File name you want to save the video as
file_name = "downloaded_video.mp4"

# Download the video
print(download_video(video_url, file_name))
