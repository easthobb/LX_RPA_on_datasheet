import os

while(1):
    print("지적도 - 1, 향측도 - 2, 현장사진 - 3")
    user_input = input()
    if(1<=int(user_input) and int(user_input)<=3):
        break
    print('재입력 요망')

file_path = "./"
filenames = os.listdir(file_path)
print(filenames)

i=1
for name in filenames:
    src = os.path.join(file_path,name)
    print(name)
    dst = name.split(".")[0] + "_" + user_input +"."+ name.split(".")[1]
    dst = os.path.join(file_path,dst)
    os.rename(src,dst)
    i+=1

print("done")
a=input()