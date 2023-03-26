import re
#---------------------------------------------------
def sanitize_name(name):
    # Remove all characters except letters, numbers, and spaces
    name = re.sub(r'[^\w\s\.]', '', name)
    # Replace spaces with underscores
    name = name.replace(' ', '_')
    # Convert to lowercase
    name = name.lower()

    name=name[0:50]
    return name

#---------------------------------------------------
def inputML(prompt):
    print(prompt)

    # Initialize the input buffer
    lines = []
    while True:
        # Read a line of input
        line = input()

        # If the line is empty, we're done
        if not line:
            break

        # Add the line to the buffer
        lines.append(line)
    return lines