import subprocess

def download_hls_video(m3u8_url, output_file, headers=None):
    try:
        # Prepare the ffmpeg command
        command = ["ffmpeg", "-i", m3u8_url, "-c", "copy", output_file]
        
        # Add headers if provided
        if headers:
            for key, value in headers.items():
                command.insert(1, "-headers")
                command.insert(2, f"{key}: {value}")
        
        print("Executing command:")
        print(" ".join(command))
        
        # Execute the command
        subprocess.run(command, check=True)
        print(f"Download completed: {output_file}")
    except subprocess.CalledProcessError as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    # URL of the .m3u8 file
    m3u8_url = "https://live.licdn.com/bitmovinneuprod/bitmovinneuprodoutputstoragecontainer/a0851f35-a001-444d-af2a-3448af3f49f5/RtmpLiveEncoding/video_1024.m3u8"
    
    # Output file name
    output_file = "downloaded_video.mp4"
    
    # Optional headers (add more if needed)
    headers = {
        "Referer": "https://www.linkedin.com/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    }
    
    # Download the video
    download_hls_video(m3u8_url, output_file, headers)