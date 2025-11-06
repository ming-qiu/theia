import os, shutil
import sys
import argparse
from moviepy.video.io.ImageSequenceClip import ImageSequenceClip
from PIL import Image, ImageDraw, ImageFont

def generate_frame_counter(begin, end, fps, output_dir, font_path):
    """Generate frame counter images and video for specified FPS."""
    
    # Create FPS-specific subdirectory
    fps_label = str(fps).replace('.', '_')
    frame_output_path = os.path.join(output_dir, f'frames_{fps_label}fps')
    
    if not os.path.exists(frame_output_path):
        os.makedirs(frame_output_path)
    
    print(f"Generating frames for {fps} fps...")
    
    # Generate frame images
    font = ImageFont.truetype(font_path, 64)
    
    for f in range(begin, end + 1):
        im = Image.new(mode="RGB", size=(200, 100))
        im_with_number = ImageDraw.Draw(im)
        im_with_number.text((10, 10), str(f).zfill(4), font=font, fill=(255, 255, 255))
        file_name = str(f).zfill(4) + '.png'
        file_path = os.path.join(frame_output_path, file_name)
        im.save(file_path)
    
    # Generate video
    print(f"Creating video for {fps} fps from frame {begin} to frame {end}...")
    clip = ImageSequenceClip(frame_output_path, fps=fps)
    video_path = os.path.join(output_dir, f'frame_counter_{fps_label}fps.mp4')
    clip.write_videofile(video_path, fps=fps)
    
    print(f"Completed video: {video_path}\n")

    if os.path.isdir(frame_output_path):
        shutil.rmtree(frame_output_path)

def main():
    parser = argparse.ArgumentParser(description='Generate frame counter videos at multiple frame rates')
    parser.add_argument('-begin', type=int, default = 0, help='Starting frame number, default is 0')
    parser.add_argument('-end', type=int, default = 4000, help='Ending frame number, default is 4000')
    parser.add_argument('-fps', type=float, default = 24, help='Frame rate, default is 24')
    parser.add_argument('-dest', type=str, default='./frame-counters', 
                        help='Output directory (default: output)')
    parser.add_argument('-font', type=str, default='./SF-Pro-Text-Regular.otf',
                        help='Path to font file (default: ./SF-Pro-Text-Regular.otf)')
    
    args = parser.parse_args()
    
    # Create main output directory
    if not os.path.exists(args.dest):
        os.makedirs(args.dest)
    
    # Generate frame counters for each FPS
    fps = args.fps
    
    generate_frame_counter(
        args.begin,
        args.end,
        fps,
        args.dest,
        args.font
    )
    
    print("Frame counter video generated successfully!")

if __name__ == "__main__":
    main()