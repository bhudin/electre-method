import xlrd
import openpyxl
import pandas as pd

df = pd.read_excel('D:\coba.xlsx', sheet_name=0)
dataset = df.values.tolist()
dummy = [0]*len(dataset[0])

weight = [0]*len(dataset[0])
bbt = 100
for i in range(len(dataset[0])):
    print('Masukkan bobot fitur dari ', df.columns[i], '(skala 1-100)')
    weight[i] = int(input())
    bbt = bbt - weight[i]
    print("Jatah bobot Anda tinggal " + str(bbt) + " poin.")
    weight[i] = weight[i]/100
    if (bbt < 0):
        raise Exception("Maaf, jatah bobot Anda telah habis. Mohon sesuaikan")
if (bbt > 0):
    raise Exception("Maaf, jatah bobot Anda belum habis. Mohon sesuaikan")

label = str(input('Jurusan apa yang Anda inginkan sesuai bobot yang dimasukkan?'))

dataset_baru = [0]*len(dataset)*(len(dummy))
dataset_baru = [dataset_baru[x:x+len(dummy)] for x in range(0, len(dataset)*len(dummy), len(dummy))]
dataset_baru2 = [0]*len(dataset)*(len(dummy))
dataset_baru2 = [dataset_baru2[x:x+len(dummy)] for x in range(0, len(dataset)*len(dummy), len(dummy))]

for i in range(len(dataset_baru)):
    for j in range(len(dummy)):
        dataset_baru[i][j] = dataset[i][j]**2
    
df=pd.DataFrame(dataset)
print("Fitur Nilai Tes Psikologi :")
print("A = IQ")
print("B = Kreativitas")
print("C = Komitmen")
print("D = Verbal")
print("E = Numerikal")
print("F = Skolastik")
print("G = Penalaran")
print("H = Ketelitian")
print("I = Sosial")
#df.columns = ["A", "B", "C", "D","E","F","G","H","I"]
print("Berikut ini adalah dataset awal dari SPK Pemilihan Jurusan Berdasarkan Tes Psikologi")
print(df)

for j in range(len(dummy)):
    for i in range(len(dataset_baru)):
        dummy[j] += dataset_baru[i][j]

print(" ")        
db=pd.DataFrame(dataset_baru)
#db.columns = ["A", "B", "C", "D","E","F","G","H","I"]
print("Kemudian setiap nilai dari instance dilakukan kuadrat dan dijumlahkan berdasarkan fiturnya.")
print("Jumlah dari kuadrat dari nilai setiap fitur kemudian di-akar yang kemudian menjadi acuan dalam normalisasi dataset awal.")
print(db)
print(dummy)
print(" ")         
        
for i in range(len(dummy)):
    dummy[i] = round(dummy[i]**(1/2),2)
print("Akar dari jumlah kuadrat setiap fitur = ", dummy)
print(" ")
    
for j in range(len(dummy)):
    for i in range(len(dataset_baru2)):
        dataset_baru2[i][j] = round(dataset[i][j]/dummy[j],2)
        
db2=pd.DataFrame(dataset_baru2)
#db2.columns = ["A", "B", "C", "D","E","F","G","H","I"]
print("Berikut ini adalah hasil normalisasi dataset.")
print(db2)
print(" ") 

for j in range(len(weight)):
    for i in range(len(dataset_baru2)):
        dataset_baru2[i][j] = round(dataset_baru2[i][j]*weight[j],3)

print("Langkah selanjutnya adalah pembobotan terhadap fitur. Hal ini dilakukan oleh pakar psikolog yang mengerti tentang kebutuhan setiap jurusan.")
print("Bobot yang diberikan untuk masuk jurusan " + label + " adalah = ", weight)
print("Berikut ini adalah hasil perkalian dengan bobot terhadap dataset yang telah dinormalisasi.")
db2_norm=pd.DataFrame(dataset_baru2)
#db2_norm.columns = ["A", "B", "C", "D","E","F","G","H","I"]
print(db2_norm)
print(" ") 
        
list_tmp = []
word = 0

dataset_baru3 = [[None]*len(dataset)]*len(dataset)
for a in range(len(dataset_baru2)):
    for b in range(len(dataset_baru2)):
        for j in range(len(weight)):
            if(dataset_baru[a][j] <= dataset_baru[b][j]):
                word = round(word + 0, 3)
            elif(dataset_baru[a][j] > dataset_baru[b][j]):
                word = round(word + weight[j], 5)
        list_tmp.append(word)
        word = 0

concordance = [list_tmp[x:x+len(dataset)] for x in range(0, len(dataset)*len(dataset), len(dataset))]

print(" ")
con=pd.DataFrame(concordance)
print("Langkah selanjutnya adalah CONCORDANCE SET. Hal ini dilakukan untuk menerapkan konsep outranking, yakni membandingkan")
print("selisih antar instance dari dataset yg telah dinormalisasi. Bagi instance yang memiliki nilai fitur terbesar maka akan dimasukkan ke CONCORDANCE SET sesuai bobotnya.")
print(con)
print(" ") 

dataset_baru4 = [[None]*len(dataset)]*len(dataset)

concordance_sum = [0]*len(concordance)
for i in range(len(concordance)):
    for j in range(len(concordance)):
        concordance_sum[i] = round(concordance_sum[i]+concordance[j][i],3)
print(concordance_sum)

total_not_zero = 0
for i in range(len(list_tmp)):
    if(list_tmp[i]!=0):
        total_not_zero+=1
C_bar = round(sum(concordance_sum)/total_not_zero,3)
print(" ")
print("Setelah melakukan CONCORDANCE SET, selanjutnya adalah menghitung nilai C Bar dengan rumus total dari nilai CONCORDANCE SET dibagi dengan jumlah nilai yang tidak sama dengan 0. Hasilnya adalah = ", C_bar)
print(" ")

dataset_baru5 = [0]*len(dataset)*len(dataset)
o = 0
for i in range(len(concordance)):
    for j in range(len(concordance)):
        if(concordance[i][j]>C_bar):
            dataset_baru5[o] = 1
            o += 1
        elif(concordance[i][j]<=C_bar):
            dataset_baru5[o] = 0
            o += 1
dataset_baru5 = [dataset_baru5[x:x+len(dataset)] for x in range(0, len(dataset)*len(dataset), len(dataset))]

con_norm = pd.DataFrame(dataset_baru5)
print("Langkah terakhir dari CONCORDANCE SET yakni menormalisasi dalam bentuk 0 dan 1. Jika nilai CONCORDANCE SET dibawah atau sama dengan C BAR maka bernilai 0 dan jika nilai CONCORDANCE SET diatas C BAR maka bernilai 1")
print(con_norm)
print(" ")

print("Berikutnya adalah DISCORDANCE SET yang menerapkan konsep outranking terhadap selisih keseluruhan instance ")
    
x = len(dataset_baru2)
for i in range(1, x+1):
    command = ""
    command = "disc_set" + str(i) + " = []"
    exec(command)
number = ""    
for a in range(len(dataset_baru2)):
    for b in range(len(dataset_baru2)):
        if(a!=b):
            for i in range(len(dummy)):
                number += str(round(dataset_baru2[a][i]-dataset_baru2[b][i],3))+" "
    command = ""
    command = "disc_set" + str(a+1) + " = '" + number + "'"
    exec(command)
    number = ""

x = len(dataset_baru2)
for i in range(1, x+1):
    command = ""
    command = "disc_set" + str(i) + "= disc_set" + str(i) + ".split()"
    exec(command)
    command = ""
    command = "disc_set" + str(i) + "= [disc_set" + str(i) + "[x:x+len(dummy)] for x in range(0, (len(dataset)-1)*len(dummy), len(dummy))]"
    exec(command)
    
for a in range(1,len(dataset)+1):    
    for i in range(len(disc_set1)):
        for j in range(len(dummy)):
            command = ""
            command = "disc_set" + str(a) + "[" + str(i) + "][" + str(j) + "] = round(float(disc_set" + str(a) + "[" + str(i) + "][" + str(j) + "]),3)"
            exec(command)
    #command = ""
    #command = "print('Berikut ini adalah selisih nilai dari siswa ke-" + str(a-1) + " terhadap semua siswa.')"
    #exec(command)
    #command = ""
    #command = "table = pd.DataFrame(disc_set" + str(a) + ")"
    #exec(command)
    #command = ""
    #command = "table.columns = ['A', 'B', 'C', 'D', 'E','F','G','H','I']"
    #exec(command)
    #command = ""
    #command = "print(table)"
    #exec(command)
    #print(" ")

print("Dari hasil selisih nilai antar instance tersebut, kemudian menghitung DISCORDANCE SET berupa matriks seperti berikut :")
for a in range(1,len(dataset)+1):    
    for b in range(len(disc_set1)):
        command = ""
        command = "disc_set" + str(a) + "_" + str(b) + " = [abs(ele) for ele in disc_set" + str(a) + "[" + str(b) + "]]"
        exec(command)

discordance_set = [0]*len(dataset)*(len(dataset))
discordance_set = [discordance_set[x:x+(len(dataset)-1)] for x in range(0, len(dataset)*len(dataset), len(dataset))]

for a in range(len(discordance_set)):
    for b in range(len(discordance_set[a])):
        try:
            command = ""
            command = "discordance_set[" + str(a) + "][" + str(b) + "] = round(abs(max(disc_set" + str(a+1) + "[" + str(b) + "])/max(disc_set" + str(a+1) + "_" + str(b) + ")),3)"
            exec(command)
        except ZeroDivisionError:
            print("Tidak bisa dibagi 0")
    
print(" ")
for i in range(len(discordance_set)):
    discordance_set[i].insert(i,0)

discord = pd.DataFrame(discordance_set)
print(discord)
    
d_sum = [0]*len(discordance_set)
for i in range(len(discordance_set)):
    for j in range(len(discordance_set)):
        d_sum[i] = round(d_sum[i]+discordance_set[j][i],3)
    
total_not_zero2 = 0
for i in range(len(discordance_set)):
    for j in range(len(discordance_set)):
        if(discordance_set[i][j]!=0):
            total_not_zero2+=1
D_bar = round(sum(d_sum)/total_not_zero2,3) 
print(" ")
print("Setelah melakukan DISCORDANCE SET, selanjutnya adalah menghitung nilai D Bar dengan rumus total dari nilai DISCORDANCE SET dibagi dengan jumlah nilai yang tidak sama dengan 0. Hasilnya adalah = ", D_bar)
print(" ")

dataset_baru6 = [0]*len(dataset)*len(dataset)
p = 0
for i in range(len(discordance_set)):
    for j in range(len(discordance_set)):
        if(discordance_set[i][j]>D_bar):
            dataset_baru6[p] = 1
            p += 1
        elif(discordance_set[i][j]<=D_bar):
            dataset_baru6[p] = 0
            p += 1
dataset_baru6 = [dataset_baru6[x:x+len(dataset)] for x in range(0, len(dataset)*len(dataset), len(dataset))]
dis_norm = pd.DataFrame(dataset_baru6)
print(dis_norm)
    
dataset_baru7 = [0]*len(dataset)*len(dataset)
k = 0
for i in range(len(dataset_baru5)):
    for j in range(len(dataset_baru5)):
        dataset_baru7[k] = dataset_baru5[i][j] and dataset_baru6[i][j]
        k += 1
print(" ")
print("Langkah terakhir dari proses ELECTRE adalah Agregasi dari CONCORDANCE SET dan DISCORDANCE SET sehingga dihasilkan matrik sebagai berikut :")
dataset_baru7 = [dataset_baru7[x:x+len(dataset)] for x in range(0, len(dataset)*len(dataset), len(dataset))]
result_norm = pd.DataFrame(dataset_baru7)
print(result_norm)

result_score = [0]*len(dataset)
for i in range(len(result_score)):
    result_score[i] = dataset[i]
    result_score[i].append(sum(dataset_baru7[i]))
result_s = pd.DataFrame(result_score)
coba_list = weight
coba_list.append("Score")
result_s.columns = coba_list
result_s.to_excel("D:\out_coba.xlsx")
print(result_s)
