from oracle import *
import sys

if len(sys.argv) < 2:
    print("Usage: python sample.py <filename>")
    sys.exit(-1)

f = open(sys.argv[1])
data = f.read()
f.close()

ctext = [(int(data[i:i+2],16)) for i in range(0, len(data), 2)]
otext = [(int(data[i:i+2],16)) for i in range(0, len(data), 2)]

pt = []

# #forr i in range(0,256,1):
# Oracle_Connect()
# #otext[15]=i
# rc = Oracle_Send(otext,3)
# print(rc)
# rc = Oracle_Send([159, 11, 19, 148, 72, 65, 168, 50, 178, 66, 27, 158, 175, 109, 152, 23, 129, 62, 201, 217, 68, 165, 200, 52, 122, 124, 166, 154, 163, 77, 141, 192],3)
# print(otext[0:32])
# print(rc)
# # if(rc == 49):
    # # break
# Oracle_Disconnect()

#0x0B padding
# for i in range(16,33):
    # otext[i] = otext[i] + 1
    # Oracle_Connect()
    # rc = Oracle_Send(otext,3)
    # if(rc==48):
        # print(32-i)
        # Oracle_Disconnect()
        # break
    # Oracle_Disconnect()
    

for i in range(48, 16, -16):
    
    tbytes = [0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00]
    k = ctext[(i-32):(i-16)]
    o = ctext[(i-32):(i-16)]
    c = ctext[(i-16):i]
    for j in range(15,-1,-1):
        Oracle_Connect()
        for x in range(15,(j-1),-1):
            k[x] = tbytes[x] ^ (16-j)
        for y in range(0,256,1):       
            if(o[j] == y and j != 5 and i == 48):
                continue
            k[j] = y
            
            if(i == 48):
                rc = Oracle_Send(ctext[0:16]+k+c,3)
            else:
                rc = Oracle_Send(k+c,2)
            if(rc == 49):
                tbytes[j] = y ^ (16-j)
                print(tbytes)
                break
        Oracle_Disconnect()
    for j in range(0, 16):
        tbytes[j] = tbytes[j] ^ o[j]
    tbytes.reverse()
    [pt.append(x) for x in tbytes]
    
    
pt.reverse()
for i in range(0, len(pt)):
    pt[i] = chr(pt[i])
    
print("".join(pt[0:-(ord(pt[-1]))]))