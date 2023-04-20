# Function to read the parameters from the txt file
def read_params_from_txt_file(file_path):
    params = {}
    with open(file_path, 'r') as f:
        for line in f:
            if line.strip():
                key, value = line.strip().split(" = ", 1)
                params[key.strip()] = value.strip()
    return params