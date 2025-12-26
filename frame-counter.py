import os, shutil
import sys
import argparse
from moviepy.video.io.ImageSequenceClip import ImageSequenceClip
from PIL import Image, ImageDraw, ImageFont
import subprocess
import timecode


def parse_args():
    parser = argparse.ArgumentParser(description='Generate frame counter videos at multiple frame rates')
    parser.add_argument('--w', type=int, default = 200, help='Width of video, default is 160')
    parser.add_argument('--h', type=int, default = 100, help='Height of video, default is 80')
    parser.add_argument('--begin', type=int, default = 1009, help='Beginning frame number, default is 1009')
    parser.add_argument('--end', type=int, default = 2000, help='Ending frame number, default is 2000')
    parser.add_argument('--fps', type=float, default = 24, help='Frame rate, default is 24')
    parser.add_argument('--dest', type=str, default='./frame-counters', 
                        help='Output directory (default: ./frame-counters)')
    parser.add_argument('--font', type=str, default='./SF-Pro-Text-Regular.otf',
                        help='Path to font file (default: ./SF-Pro-Text-Regular.otf)')

    return parser.parse_args()

def generate_frame_counter(w, h, begin, end, fps, output_dir, font_path):
    """Generate frame counter images and video for specified FPS."""
    
    # Create FPS-specific subdirectory
    fps_label = str(fps).replace('.', '_').replace('_0', '')
    frame_output_path = os.path.join(output_dir, f'frames_{fps_label}fps')
    
    if not os.path.exists(frame_output_path):
        os.makedirs(frame_output_path)
    
    print(f"Generating frames for {fps} fps...")
    
    # Generate frame images
    font = ImageFont.truetype(font_path, int(0.75 * h))
    
    for f in range(begin, end + 1):
        im = Image.new(mode="RGB", size=(w, h))
        im_with_number = ImageDraw.Draw(im)
        im_with_number.text((int(0.1 * h), int(0.1 * h)), str(f), font=font, fill=(255, 255, 255), align = 'center')
        file_name = str(f).zfill(4) + '.png'
        file_path = os.path.join(frame_output_path, file_name)
        im.save(file_path)
    
    # Generate video
    print(f"Creating video for {fps} fps from frame {begin} to frame {end}...")
    clip = ImageSequenceClip(frame_output_path, fps=fps)
    
    # Create temporary video without timecode
    temp_video_path = os.path.join(output_dir, f'temp_{fps_label}fps.mp4')
    clip.write_videofile(temp_video_path, fps=float(fps))
    
    # Calculate starting timecode using timecode package
    tc = timecode.Timecode(fps, frames=begin) + 1
    start_timecode = str(tc)
    print(f"Setting starting timecode to: {start_timecode}")
    
    # Add timecode metadata using ffmpeg
    video_path = os.path.join(output_dir, f'frame_counter_{fps_label}fps.mp4')
    
    ffmpeg_cmd = [
        'ffmpeg', '-i', temp_video_path,
        '-c', 'copy',  # Copy streams without re-encoding
        '-timecode', start_timecode,
        '-y',  # Overwrite output file
        video_path
    ]
    
    try:
        subprocess.run(ffmpeg_cmd, check=True, capture_output=True)
        print(f"Completed video with timecode metadata: {video_path}\n")
        # Remove temporary file
        os.remove(temp_video_path)
    except subprocess.CalledProcessError as e:
        print(f"Warning: Failed to add timecode metadata. Using video without timecode.")
        print(f"Error: {e.stderr.decode()}")
        # If ffmpeg fails, just rename the temp file
        os.rename(temp_video_path, video_path)
    
    # Clean up frame directory
    if os.path.isdir(frame_output_path):
        shutil.rmtree(frame_output_path)

if __name__ == "__main__":
    args = parse_args()
    
    # Create main output directory
    if not os.path.exists(args.dest):
        os.makedirs(args.dest)

    if args.w < 100 or args.h < 50:
        print("ERROR: video size needs to be at least 100 x 50")
        sys.exit(1)
    
    # Generate frame counter
    generate_frame_counter(
        args.w,
        args.h,
        args.begin,
        args.end,
        args.fps,
        args.dest,
        args.font
    )
    
    print("Frame counter video generated successfully!")