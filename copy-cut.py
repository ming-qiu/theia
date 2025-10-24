"""
DaVinci Resolve - Copy cut points from track 1 to track 2
Creates new clips on track 3 from track 2's media with track 1's cut points.
"""

import DaVinciResolveScript as dvr

# Connect to Resolve
resolve = dvr.scriptapp("Resolve")
project = resolve.GetProjectManager().GetCurrentProject()
timeline = project.GetCurrentTimeline()
mediapool = project.GetMediaPool()

print(f"Timeline: {timeline.GetName()}")

# Get clips from both tracks
track1_clips = timeline.GetItemListInTrack("video", 1)
track2_clips = timeline.GetItemListInTrack("video", 2)

print(f"Track 1: {len(track1_clips)} clips")
print(f"Track 2: {len(track2_clips)} clips")

# Ensure track 3 exists
if timeline.GetTrackCount("video") < 3:
    print("Creating video track 3...")
    timeline.AddTrack("video", None)

# Get cut points from track 1
cut_points = [clip.GetStart() for clip in track1_clips]
cut_points.append(track1_clips[-1].GetEnd())  # Add final end point

print(f"Found {len(cut_points)-1} segments to create")

# Assume track 2 has one long clip - get its media pool item
if len(track2_clips) != 1:
    print(f"Warning: Track 2 has {len(track2_clips)} clips. Expected 1 long clip.")
    print("Using first clip as source.")

source_clip = track2_clips[0]
media_item = source_clip.GetMediaPoolItem()
source_start_frame = source_clip.GetSourceStartFrame()
source_offset = source_clip.GetLeftOffset()

print(f"Source media: {media_item.GetName()}")

# Create clip info for each segment
clips_to_add = []
for i in range(len(cut_points) - 1):
    record_in = cut_points[i]
    record_out = cut_points[i + 1]
    segment_duration = record_out - record_in
    
    # Calculate source in/out points
    # Source frame = where we are in the original media
    source_in = source_start_frame + source_offset + (record_in - source_clip.GetStart())
    source_out = source_in + segment_duration
    
    clips_to_add.append({
        "mediaPoolItem": media_item,
        "startFrame": int(source_in),
        "endFrame": int(source_out),  # endFrame is inclusive
        "trackIndex": 3,
        "recordFrame": int(record_in)
    })
    
    print(f"Segment {i+1}: Record {record_in}-{record_out}, Source {source_in}-{source_out}")

# Add all clips to track 3
print("\nAdding clips to track 3...")
result = mediapool.AppendToTimeline(clips_to_add)

if result:
    print(f"Success! Added {len(result)} clips to track 3")
else:
    print("Error: Failed to add clips to timeline")

print("\nDone! Track 3 now has the same cut structure as track 1.")