import os
import re
import pandas as pd

class FetchCode:
    def __init__(self):
        self.path = "C:\\Users\\dmusalappagari\\Downloads\\Baker\\Baker"
        self.file_name = ''
        self.pattern = '(?:PD|IB)\d{2,}[A-Z]{2,3}'
        self.tags = set()
        self.tags_code = {}
        self.tags_obj = {}

    def read_folder(self):
        files = os.listdir(self.path)
        return files

    def find_tags(self, content):
        matches = re.findall(self.pattern, content, re.DOTALL)
        self.tags.update(matches)
    
    def open_file_content(self, file_path):
        with open(file_path, 'r', encoding='latin-1') as file:
            content = file.read()
        self.find_objects(content)
    
    def find_tags_obj(self, line):
        if 'OBJECT' in line:
            words = line.split(maxsplit=3)
            return words[1:]
        else:
            words = line.split(maxsplit=2)
            return words


    def find_text_between_tags(self, content):
        with open('test.txt', 'w', encoding='utf-8') as file:
            file.write(content)
        for tag in self.tags:
            matches = re.findall(f'<{tag}>.*?</{tag}>', content, re.DOTALL)
            if matches:
                if tag in self.tags_code:
                    self.tags_code[tag].extend(matches)
                else:
                    self.tags_code[tag] = matches
                self.tags_code[tag].append(self.find_tags_obj(content.splitlines()[0]))
            else:
                # find the line that contains the tag and add the entire line to the dictionary
                match1 = re.findall(f'^.*{tag}.*$', content, re.MULTILINE)
                if match1:
                    if tag in self.tags_code:
                        self.tags_code[tag].extend(match1)
                    else:
                        self.tags_code[tag] = match1
                    self.tags_code[tag].append(self.find_tags_obj(content.splitlines()[0]))



    def findObj(self, content, exp):
        ep = r'(?<=})'+f"\s+(?=OBJECT {exp} \d)"
        blocks = re.split(ep, content.strip())
        # take first block and find the tags in it and then find the text between tags
        for block in blocks:
            self.find_tags(block)
            self.find_text_between_tags(block)

    

    def find_objects(self, content):
        # check if first line contains pageextraction
        file_types = {'codeunit': 'Codeunit', 'page': 'Page', 'report': 'Report', 'table': 'Table', 'query': 'Query', 'xmlport': 'Xmlport', 'pageextension': 'PageExtension', 'tableextension': 'TableExtension', 'enum': 'Enum', 'interface': 'Interface', 'permissionset': 'PermissionSet', 'enumextension': 'EnumExtension'}

        for key, value in file_types.items():
            if key in self.file_name.lower():
                self.findObj(content, value)
                break



if __name__ == "__main__":
    fetch = FetchCode()
    files = fetch.read_folder()
    for file in files:
        file_path = os.path.join(fetch.path, file)
        fetch.file_name = file
        print(f'Processing file: {file_path}')
        if os.path.isfile(file_path):
            fetch.open_file_content(file_path)
            
    # print(fetch.tags_code)
    result = {}
    for key, value in fetch.tags_code.items():
        strings = [item for item in value if isinstance(item, str)]
        lists = [item for item in value if isinstance(item, list)]

        joined_lists = [', '.join(items) for items in zip(*lists)]
        result[key] = [strings] + joined_lists
        # if key == 'PD34120RNC':
        #     print(strings)
        #     print(lists)
        #     print(joined_lists)
        #     print(result[key])
        #     break

    rows = []
    row_names = ['Code', 'Object Type', 'Object ID', 'Object Name']
    for tag, values in result.items():
        row = {'tag': tag}
        row.update({f'{row_names[i]}': value for i, value in enumerate(values)})
        rows.append(row)

    df = pd.DataFrame(rows)
    df.to_excel('baker.xlsx', index=False)
