from pathlib import Path

input_file = "templates/index.html"
output_file = "templates/index_cleaned.html"

with open(input_file, "r", encoding="utf-8") as f:
    content = f.read()

# Fix escaped HTML tags
content = content.replace("\\<", "<")
content = content.replace("\\>", ">")

# Fix escaped css star
content = content.replace("\\*", "*")

# Fix escaped underscores
content = content.replace("\\_", "_")

with open(output_file, "w", encoding="utf-8") as f:
    f.write(content)

print(f"Created: {output_file}")
