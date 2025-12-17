import os
import imageio

gif_path = "static/1.gif"
out_dir = "static/actor"

os.makedirs(out_dir, exist_ok=True)

gif = imageio.mimread(gif_path)

for i, frame in enumerate(gif):
    imageio.imwrite(
        os.path.join(out_dir, f"frame_{i:04d}.png"),
        frame
    )

print("PNG 序列生成完成")
