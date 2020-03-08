#!/usr/bin/python 

key = [0xF2,0x1A,0x04,0x9B,0xD0,0x73,0x23,0xC8,0x39,0x98,0xCE,0x09,0x0E,0xBC,0x86,0xDA,0xC9,0xE0,0x39,0x89,0x2A,0x5F,0x72,0x67,0x83,0xA5,0x61,0xFD,0x25,0xEE,0x30] 

cs = []
cs.append('BB3A65F6F0034FA957F6A767699CE7FABA855AFB4F2B520AEAD612944A801E')
cs.append('BA7F24F2A35357A05CB8A16762C5A6AAAC924AE6447F0608A3D11388569A1E')
cs.append('A67261BBB30651BA5CF6BA297ED0E7B4E9894AA95E300247F0C0028F409A1E')
cs.append('A57261F5F0004BA74CF4AA2979D9A6B7AC854DA95E305203EC8515954C9D0F')
cs.append('BB3A70F3B91D48E84DF0AB702ECFEEB5BC8C5DA94C301E0BECD241954C831E')
cs.append('A6726DE8F01A50E849EDBC6C7C9CF2B2A88E19FD423E0647ECCB04DD4C9D1E')
cs.append('BC7570BBBF1D46E85AF9AA6C7A9CEFA9E9825CFD5E3A0047F7CD009305A71E')

c = []
s = []
for i in range(0,7):
    sp = []
    for j in range(0,31):
        sp.append(0x00)
    s.append(sp)
    
for i in cs:
    ch = []
    for j in range(0,len(i),2):
       ch.append(int(i[j] + i[j+1],16))
    c.append(ch)
    
for i in range(0,len(c)):
    a = i + 1
    b = i + 2
    if (a) == len(c):
        a = 0
        b = 1
    if (a) == (len(c)-1):
        b = 0
    for j in range(0,len(c[i])):
        if((c[i][j] ^ c[a][j]) >= 64 and (c[i][j] ^ c[a][j] <= 127)) and ((c[i][j] ^ c[b][j]) >= 64 and (c[i][j] ^ c[b][j] <= 127)):
            s[i][j] = 0x20
            if(key[j] == 0x00):
                key[j] = c[i][j] ^ 0x20
                
for i in range(0,len(c)):
    st = []
    for j in range(0, len(c[i])):
        if(key[j] == 0x00):
            st.append("-")
        else:
            st.append(chr(c[i][j] ^ key[j]))
    print("".join(st))
        