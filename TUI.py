#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# OpenSrc By ziphercyprex
"""
Thai ID + Thai phone structure decoder + automatic BORA RCODE updater.

คุณสมบัติหลัก
-------------
- ไฟล์ Python ไฟล์เดียว ใช้เฉพาะ Python Standard Library
- ทุกครั้งที่รันจะลองอัปเดตฐาน RCODE จากไฟล์ทางการของ BORA/DOPA
- ใช้ HTTP ETag / Last-Modified เมื่อเซิร์ฟเวอร์รองรับ
- ถ้าอัปเดตไม่ได้ จะใช้ rcode.json เดิมต่อโดยไม่ทำให้โปรแกรมหยุด
- ถ้ายังไม่มี rcode.json และดาวน์โหลดไม่ได้ จะสร้างฐาน fallback ให้เอง
- รับเลขเดียวแล้วแยกอัตโนมัติระหว่างบัตรประชาชนไทยกับเบอร์โทรไทย
- ตรวจรูปแบบเลข 13 หลักและ checksum
- รองรับมือถือ/VoIP/โทรศัพท์ประจำที่/+66/เลขสั้น พร้อมถอด prefix/รหัสพื้นที่
- แสดง historical carrier hint โดยไม่อ้างว่าเป็นค่ายปัจจุบัน (รองรับข้อเท็จจริงเรื่อง MNP)
- lookup รหัสสำนักทะเบียนเต็ม 4 หลัก
- แสดงประเภทสำนักงาน จังหวัด ชื่อพื้นที่ สถานะรหัส และวันที่ยกเลิก
- แยกหลัก 6-10 และ 11-12 โดยไม่ตีความเกินข้อมูลที่ยืนยันได้
- รองรับ --lookup, --find, --compare, --json, --verbose และ --self-test

ข้อจำกัด
--------
โปรแกรมตรวจได้เพียงโครงสร้างของตัวเลขและข้อมูลอ้างอิงสาธารณะเท่านั้น
ไม่สามารถยืนยันว่าเลขถูกออกให้บุคคลจริง และไม่สามารถถอดชื่อ วันเกิด เพศ
ที่อยู่ปัจจุบัน หรือความสัมพันธ์ระหว่างบุคคลได้
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import unicodedata
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
import zlib
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable


APP_NAME = "Thai ID + Phone Decoder"
APP_VERSION = "4.0.0"
CACHE_SCHEMA_VERSION = 1

SOURCE_PAGE_URL = (
    "https://stat.bora.dopa.go.th/stat/statnew/statMenu/newStat/ccaa.php"
)
SOURCE_XLSX_URL = "https://stat.bora.dopa.go.th/dload/rcode.xlsx"
DEFAULT_CACHE = Path.cwd() / "rcode.json"
DEFAULT_TIMEOUT = 12.0
MAX_DOWNLOAD_BYTES = 20 * 1024 * 1024
MAX_XLSX_UNCOMPRESSED_BYTES = 80 * 1024 * 1024

# Snapshot ฝังในไฟล์เพื่อให้การรันครั้งแรกแบบ offline ยังสร้าง rcode.json ได้
# ข้อมูลนี้ถูกบีบอัดด้วย zlib และเข้ารหัส Base85; ไม่ใช่ข้อมูลบุคคล
EMBEDDED_RCODE_ZLIB_B85 = (
    "c-ri}ZEqV%mM-"
    "{L2!9^;NPOo2_e0sfy&llg2$p&cW3ecNZMK=9$P_5@`Wk}^l+?n=c9(Z&yX02MuIp||5L7L<2exc?N&m(_Wiw7>kj%(PMkW<Op2K"
    "dSS}jW=`Q&*rBTk%`=luJ%cB980u^zkjFaLgR)9-"
    "g5Hm?24wXYXvUoSp=y*T)K@$u`$*4K;IUoVcoUVQ#~@$T!zIefn*zuEtK@!{*`cknk~zFr&!ugbx%gWnyF{M{S)eh0pP|Mg<`+CP"
    "P#_iews9(-oA*W1|qmw*1}%^vIh>Ve;7zk1|vuwOm)f7NUL^P_d2J^JVFL;q3ZSO0T;^M4Ls{9A9GBaZxvU(dH}^IQM7-"
    "}aipFBb)%k8nXp00>~_OXYI#=Z_nmMwj&(kG^HS;IH!w`SSIJ>h*<U@Vjo~q2GP<?L)uc34WKaEY|+%TF37-"
    "Hn085|8)(n<sH2785rWr*UM))_OR#ukH&xdEnDAg<iIj-zFzLaXKVf(j1<`IRN@=_{25U2j{lJV6F^`74*urF*Ncm<7kghXw!dCJ"
    "2Xh8*a{2Y*BtSeg`Q;A(Ie+b2$>8AA-^1TL19)4&>>hmLIsEfK;hTf_`}6<!PtM^@4&mL7a?!Vb%ioT_@6TT^{*r@V1i$_T-"
    "X`{bLyUifp9CNNO8(96wg3Bn{nNFLuKz#0&O>hTf4{c4z`u4_yCImBH(FqgUHHs>`20h8qaKJmE-"
    "(wcrqKyrbAxprxBS+%;Jp_Y)wdH@bpdZAZrWpMlymisR$05>%dK=-"
    "i*<tk6ENXz^Xl6kNNQihn+NatCGu9wey7)D5Bgm%`bGuyjo!e!TmZ@+<t0YnYlZci{dR7x$vVor71eip3s61`>;#-"
    "0!~OUko{<aq6#mFWKcd-Xxn<UB`kmaWm;2^vzZbxN%Ay}vQa|p!ypea|D#BG@__p749tW`f;A`Z4%If=^$U9nR$~*o>?ncaz6%Em"
    "!0m!?rm)|G8*{T<SiN9M_eYY3lNdQ9z*S8;elbcQE1=i|C-"
    "=n5MB*;j@!n*N0B)_16a+3HSH#uXm>&eF+0|Den!_i&#R3#`MyWO9SEM|#FnD&;>$8Pyo@)_KVyun&mv1mSadq0<3(~FSoaD_Nl@"
    "8<x_!=~TpJl<q}^gT<O2k<ytJNaXBka&T?y6~{+>8xGWtPbMoJCJ+RR&;d`Pug-|`d0L<DoHod<rZJF_(FQB+(fxLYeX-yu#gByT"
    "<QyW-!HJ}oa>O?Wxa<@w&}H_D?lN3Ukz;qUq8ZI2HV>A(G{AT>@I6Z-z=}W)(h#E9E8Lw@3b1cQ*>=E#O|-"
    "bl6NL=m%w&^QI>%mH(Von!xf701&$+n4Nu^}bP3D!ZypQ`*NZLHh1h);7-<_I{-|z4_t~c3<19qppq#iw{TVje!0w-"
    "8tJG7r!MBdc8&xzH`AOcY!hv`uS$G*;&fA^f5e?qA8$4RkH?3-"
    "JjK9HhEH>J|D)&S1Xa@JK@+Pqr6nrIEvfvFrf>(6{aQQ{6J1$?@S?&(M&zk+nTNh(b#5TO|fqbrlD?W(4U7ZsUJXukQ#YO!cvIho"
    "Bx-lLFZx!FYi|YN)EgxLmk+S<Wl$RC8S!qWsIC0C=`JNSdpV-YlaI(&%i!a{oF!F9otjThG>uNQX2fA2_?>eF9cLW#r9^B^Ws(;t"
    "_a{oR07#7PKpOD`}a+5y*Q(vfSL!GS$OO`^IZb97wJL_D@Rn~%axurgS#cJY``WC)F)Gebc0qMSCq?*R;7l3IGuJWD6F<t7n9tCE"
    "PLMkn2?w=R(%(e?hi+92qSJnkDaF}yT%KcDU)Ocl|;PNijha-"
    "5zf^|iKlh?ogjc^cSuV;$WQ7Y&kE#WQf0r9W+sX}4>H@#kPGd}X7_fSc_yZLht-mzj0?)f&NFD=EX4&B@Xg=ED7WsN&7HXx0|9eC"
    "H#UAt5n&u6#+9qGuttCCsOT)sG#{TF-{tA#i>9w^UuBX^sx7^})7P)_zL&R}6a4%ZBYAbG4{PZWEV6Bg5bxv?d!r#-"
    "%w#TV3aUbAj|lFSQd@hzD2SXnoMyP5GFX{D)n1ZAD`dak;2Yb@%;m5W-"
    ">hVS>=gUeIgn!rXqb%`#QG&fZ682shy#SiigPgF1~O|Q4fI{kIUwJewQZ}v<+S;y+z1m3P<4c#3YS`;?uOZeso^$y@>eaMdnQT{6"
    "nnp28!u@YQx;a$a!7?nl+hVc@vF4?0G9FXAcI+{03a82LJoBW65LtTG^FVwL^LZy)S^t^%ijD7<NTuEgEsT2(%4zYXh=z|-"
    "p@vW^=GQZ)j#2Rx4!nX|;j8?FcDp1RY`$c=qk~fQLz>5Cl_5nyfh4<T$=a4V(q<sus1@(612f04m2z(+h_>j7~u~w58y1~)nM|me"
    "44(?j*q6h0#yUo<IZFNx}_xccqU6oy`-enJ>=BVa1kMl=hYFD$_bRKseN6&KA*bNEi2ytBEe%dSLJ^C$WajX^-"
    "4{+zONc}(J0GGIs2Zy=9Eb)cBTGEFv4z%HmjvBtG>W)ex!uD2uDsBeHqYd?ms4ZwNY!_~+6V<lj&R>t(wx$b1z8XYyN1vXavMAQ_"
    "`f$ld^}^S>ss~jo=x>RAfPJC9L2y^_Am4gleVd}@LHh_eUL^)7dn$)D&4HJ57&G8~rW#KdP7(LikFZyribIO>yal&X(|Zzm%>EY-"
    "_doVFdtFaPn-?s`oB!#W|J&bs4}*7ly1`Qx!3N&wMt?d&|7i9_lVX8e`<KB%blrc**1g`-"
    "Zzq8&BkbSu912gDER6o}O7Ez%g2<0t`<Gs~-"
    "}t9%kG#!?ey8Vk`i)25J_vpsxFYMykG)M_{G7*47Zzeq%l;r^0BMBJ!YlxB8}33i&Yc*{z6I<*?8V4>AmD%dNc#r~yzPc&kwB7&("
    "!is}Nj@_tdjMLT(4%|Dt^Q`<NzR4}6%{oe3YmEr4F`#n-c~#ubh8;84`#!N9v?McikW$NEi+|eH{XiwDZwdzR-"
    "9<CQ6r=@F(Khk4&sE=n~lzl85sdbjgj)?kH{A6o^RBYh9HG)c(Z;?BKWBBQpwEAXpur#*P^RXxW{JowCDj+Bc?h5F~jnkK_dG##3"
    "???lj9}77+G)>r42$$Y$F#OLj%w}UE$4)!0Z=7jo#YC$DA;7d=Njv2PYjKzv3?4Y>zroh{ULgSD_iHIW$y@EXhid+0degM~#;3fq"
    "@fISQHJXevIcIdpRCbWixwh#^I#KQua{6k;Es8PEO-7L7s$v!e&H{9v(GXvPTZiWC7RPVeW)q{Vv~!)j}*tn(~UjC$rmp%z&wpR-"
    "D8~c*!C1!t4EbjR8-"
    "uc(d~jBG9Ojkv*jF<2aXKZBVN_8#;cZ?+{!qr0^b6o1!f0rJ7<5QReP54ABJ*RBBvRCWu@^eKTyp+cM7=_1tdv8PC{1nVsjB2$C8"
    "<*`p0R;I|G+$Ijz5v+GbeHnT909w0Shvd0O;!AJajS5}4}C^c&99^`#yMouC~V*KP6uD%KJ--eaZ+D&NLN3#h<p;F^&adNJ<0sVn"
    "QAJ%E;l0<!8-"
    "tVzlr;3shQ)4ZAMB&e}=IqP329_1x(==#cnGIXxkW%C6O2rAhE_*vx_Zs%9G}^PTP8b15jhO6Vgm}^BNImz(2ek~FpJQ2PRslLQS"
    "Zah#UcJEKo`3=O;(OCdbDi~Oy)8umQX{53Ngdp0*;%-"
    "!o?mZy!M+2%CNnFCQ^8SVq%z5)GU!FU5E31_cGP}*c29~1A2nX8ljIxDMU#nI6hFJt-3TUX6x1f)MSd@<&)>&)k=1@<)_X_<9x+-"
    "9*#id8;Om7H-q?z_t*~DIVRL5nqI!JPc)3DwU<bO{w*c`;Y(EiU0Y7(VcITLJNU8CZpQK>%SoYYQe!civ?f-"
    "ky=*_NDQAI|Lj>1(E2mFVx)UaomI9TIa7c<Wsj3A`OOmPwy>=f{c)!|e<udMc2CwHgMW?AB)0Z5IQ(j*>$sD{<)s@>#G;X+t#X8X"
    "U2jv6Hs9qixBpBzd7MCEuNoLiRq4K~a3y&fYqdMXn}FkeW}r+V>-"
    "UbFGUn`zimMM#aA>f}YaZ>1#ExhiAFD{^@?yFcrCToocUcCv>6M8C`*M73*#l>d`GYR)dmui_%cL~-"
    "&m0e)KB(VZyStm7~e88t2zCrMPENGqLb6j|=D7FTzfS&~u%jv6WX2?NGQP*D6Xu0-"
    "U4sNF2%#t~@L_?YNi|8b;I;z+ZxgSyf(%iW*Z)}9EG8b1@A?mqxlf$|KJ<jMh`fnM2af0ieFJxFT&lqPh&hT`dRWsj9T-"
    "(ylqeSLOkO9NVJbd@dCKObtE*9KM4ydfS<A@m{DKDJ_I@+?j4L<4^GGMjZyhUkeV`^i6ya*R7<t~J>cFL+Y87TwIwnu$26QB<3Rq"
    "6=A*FGMac6n^?}_U!g&$BhOcF=9$rXc~B}QE7&S?1n${lJg-pVuaMH8i`JwDPnW3AP8~-POC=8F*-"
    "2eQQPZq`?H=GaK$>0a(9IwrTD?W0U^^DlC*ZI)JU%xNY~nPOYf84Y;*!1*Hbb2Q2g!WoJKID;ne&tw76T0E4MzA=aF}E-"
    "`=w^?@Q%wL;H&{AjU9cfEJ@<F^Q54?`6K;D^s(m9Vx@_l|MuEQ6w3ejiWG|8iAg;=Z)X%BA^mT$_|w(XZi7*jc)M?WuZ5GmrUpYP"
    "_=a-t*u3?#20NNhc_K-4<S&scF|z%fjaY>k=bf{T5Yue-"
    "J|i{8Wv1H4i8lhN@+QW!VWtk_d=f~Hb<n(!kyq|v+>BA4l8DaR7onQT$cDx&ZGeUuBJbL`KD?#iz`$N4N>z~v*bVbK?+lt;3Vdyj"
    "Hc&xLn1@n+(tl?|Abelb-ah(2ADX;6xjaTUhsEXu1sLkR1gdy&X_tF>8#=I8XqKmeW7@Lv3gC%As;DS2!AiQELh|6jl3+OyYMNSy"
    "S&0$!9Swz^2|7&<Wi_U%*Bas03Pf6B!`S<UvgLb53bAAH4V#~{Rd+g<FNq$Zr0|K6*@=SZoXdpNtPA<Bt(pra-"
    "v2%_iLVuM?t6wI8|E9cAio@-YbeHcmCbnYQ^r-"
    "Y&fK&T94_U=Bri9?2QX&;o^>b+^br%!<|yia+;(&*oQ(k1QhB46J@csV4uZ3{Y2%E$XghBj6kMJ<)Y=qG`O5g%|6&iwvvW0sj`;0"
    "b{G-f-WF(QoyX^(w*vnZ&G|#cN|n8W{m4K<_bEJlGGU4>EHJ(sl?a!XqCCUBB&^%u({qkv&BS~gWbH1RD9<?9R2ePV4BTGwKNd&l"
    "Fk~Vw_Om>T9qXoWaUmanleN9h<Beby_GTKe&2R}+DYp~^iaY`E$3i>w7xi>0pW)?TWybL!8=zAqwqosBK9S{Q39=}>U0tZ6%<Ng!"
    "L!?Sw)f#n^RCmmb#5pYes9hoGb0CY4C$8eW8CW<ClMiFRN?uQg1+ITrWu5ty6}O_Zc0Y0K(X!8{ggxMMkKEVT>cPYa$R`nCc)jVW"
    "u59v*m9!??hljxtuv4s^rOrj_3RzJff2teik*jz6Y)X}z5g2sLoVGZ2+~BV9Q@d;ypN${Mh1ST$^wr*ceLdCd&WSu(pG~paPX~yO"
    "*$ri|Fc=YFCCmfjnkHjs0a;kJ^%&^b*i^+TtI+sQV2QlkRc9rvuvV~8kJE&;a7QQNAY<y%qNaWB96CkM<+PO3P$43W6z&fKM!Bo{"
    "GHbjxDtI1lbXn(NGnc`&c>u4Ny!mc;m1EKs4z7LbH^!KH*ka}SLiswJT(tGD3&rsFaHm|M95Si4nO~IP=xw5w0PIh9gNJ#gDfI+q"
    "UhtfZ8xy^m%ANDol6`u2<OCo6B%6A(pW}0fX3wVzlrW8UBH_`ZU+|AVtCbzYdga?5j^d7F)MWN!N0p7ief?TLN@@dTbz_K^4sBM0"
    "-JjtnCj)h1CCu%${myI`PCZ6CY*{K|{8|3+JW)|%m94W}eH~5tMj~V~<Xxe5QF+vc$K*qi7}scz&g>b}fuqBboub6lu@n@}2l1eb"
    "rIN(-v0Yr{Y&3(wwkr|&V)IN>91)fwqN_rT;-"
    "y2IrJ}<(atA*em?Y{euJhm;8rJl1>9FP?MDnw^3}XrKs@E35(b)$}5+TzevTUD7bHaG`%+bsyU!ewYSbdZwV>E+;Bl?LXi{*Y3C9"
    "Y@xbWzS~TdDBNP&GD#M~rsqQXysK-kVTdncuiS%Epc_$~xC;_W7(q)|p26ULvBH=?h7wH@}7X>aJ}6jAeW&hzQf&5yVrh<(n5`Of"
    "`L@9lr^+|6!A|VJNDFf{{cl#B6RV@cR+ggqPzAKRQpHSFa|(AZBY*o!>jD5I)dqinS{(i+;lfb878GAjAxA61p4=mK0AgpiPqL6J"
    "14Zq15Nn_-0VVSSXuW_yk`s;mEWVm)~j!mJg0}%(dM$Fo@Z@Vz9NGH^lo3gDq@JZPv-"
    "vn|{lwnKcM#*?4TBr7Eg&31463GT-#Pq$=Q|Co3w`JOs{5DX(~Sv*FoR-DIIsv)RIWE+-"
    "<KYTC~EX0ytk%re%fhe+#DtuENF!MfB=`&nuBJNW9aE)#0zaFb?D2PxD@NS+cSV*2ZdlA#HZi7}YBxln!@@fi|*b5K=o%~V8t98_"
    "Zb6<k%2pN^{_zlW%T{4}|OeAZ4&rZVz^-DZ!!!8b=@syVpue~fb3=ZX#Yy<T&+YW8}Z#GXPsGS`uL^zYOo-yPP<-"
    "40HCD7HThQglyWsdRl|G59a(>08W)zZV?3sZaPI#GiFzBpx9tPZ&ZZ_S9AED+0dK){?CG4G*6Z;46~nJWD#%86uG=h?|FyrL7hYu"
    "!St$)|?JJM3jxdq(zo3x#@wN);sk`ewQ^t^$P`ADokQzEn0IR)2Ah|IFj>_zpUAdqnw4YfAROjw)elNgHTpkhjpM$e`ao^0V^>ME"
    "%kL|sKY?<2+l*nM$KOABx~u??4G_kDQv2hET?`FSc9Y79fbW<l)1APHIaHhJOsE67TJyI7S-"
    "2hT6V67NQ|+H&6P4NE6~?N%yMBI0wYFG)k0Iiv8)t71I>t6%CQYv+h)>oqX9;Yo|?@$2bsDCW;xAZ4r<nmDVsS4QsRTBN%mfXx&U"
    "0okMK==@K0`;wVHk>c)D`mJneJkw5Mze-"
    "5%K>$5<;Su{PoRX0yqaX*KkS*eBFw$>rZSm8Pzef&cX#fRDwtxQlq#b;Z3N#Jo{qYBm{wPdIFa_}q8*=z+`|9cgu^wON&L()pB>y"
    "~iX>h?7vtApU(W4*qKEH0w-F?K?y;ZzPn{lJEwu<-JBkQ<MR-"
    ">)w>eP?5|VC6#QHyaR7mOG%SWm69Qnd84Fipya18Q^Qy5cdA<0N|TqGIo>Zc2<DB1nt_Br3l-1UslUk{Hr*OO^2u{%_Qi#?t)v6-"
    "kTKKgxY&t%C)NOI-WW0+*-"
    "vHcBSC`lo|stR_8HNGm^Ui&W`&J2&4hr}Mu)l9%{WU|_$O}kw4<6XelNR>j&ujQX${1RpqMu<OegF2a*p6fm4)8~VB9jGe)hcC-g"
    "KZognjT_ugSi|c(jkX9X6yvFmEK757dKB$IC5wQCsSIl;c<%EyoK|0+KnSq-<03G}5Metm)`$b^S-"
    "PY(&ukoHs)4jY=0CtI5K(To@$OIKwg`T@QKQD6ZJ0m_H2f<B~5vE><PVc_XuGxgzmkAfpra<ClGq@<CRU7(qO5^wuoXyCtvs42t2"
    "OYopusE}y%IO?pHRea;B4ELa=R_Cks<?T*ZpPY6@pY<hB#BAS6IsV-sO=(p5eHCk`~46=D!y0(CTb|hW;dcgC>Zr<8l5Aigd$*Hq"
    "C1@x6J+vMUWO&`&Kk3`(_Ms>lu-"
    "c*W^1~$E^xsH8Q>`dw5p6fBZ<b^bx(!4BIYAomE1ibS`uBA>dV_dk)MHw+Xw(}o9u2_m~ner?s=$!L}OvUyf<nj4SWSTm2o#(7nt"
    "tm`EfeQb#)4@S9!tU)p%iZ(PM4%H9&l|%v%Tk`NWS7I4PF!h?N3&+G;i+PrGfr)(B2QTIhDC<y@2ED;mZb9d+ENt6&T|oJHE;RBe"
    "F6Gb+iL!vKd-?l_K}b5S&hJ*H*W1|E~hKQpV=Ne|26tmOGBXt7aMl@$g%Wdfzvb=|7#})-KxQIj6{jed<pZ$eaZgxf-"
    "{&S`&>W3fTw>EjxSI#&l}BU%R!A)jfJ=(8*|=)Xj=(7U85pKis!bY(UxxXARzx%Zf@`3hd-"
    "$5nowTOudIzMlb{DaZ*<!((Hh&5kLyXE+unVPlGMT}-I3oPDhg-"
    "4z^iWId3N=H=Z)P(%i|cvN}1rA#xS><C$+q_?C!CY8GflLo%^QCI#02b&FR6;8?%;bsWJ&cp`iEgLNTglDX=Bg&XVZmK`cbv>fz2"
    "CxAyu2-$Ri8NUwnS%~QX}q0cP-tU)(#RF<uECOivcc;WZEYO3gFlli%u%|;JRBZ40ByfIv{yAT_-FHx<-"
    "b1jm21+8vn9Ud#QGOOixt~FS$Amlx~F6+|5GmvEedsZASrPKs}FTKiLeN)gkd?*EU1i2Uc)?<nBQcB{*@Or~VUD@D^X1yuJKOJlN"
    "Q0ica)v+d(85<cUOv+z7(2KGiZ)#=UQ{SnbBcDvho+eSRHzUJi4%PVA3d=3|Zq;)%pHIfN29Yj)NpE3033l!86B;-"
    "9oO9<(L?RF}rZ#AD@#e!b5m*rkNkyaN35*22?N%I6gMo~F%VzD%&+yNm`4~l>H2Mf$(hI=+GOB~=Z8o6((}Pw<*K(K!Tx2LP=~WL"
    "ji?*To@wqC)({j4~={}x{X)RT0gv)4Ph!5%-zRA$^>kbi4o1WjY^-Uc=*Yc%O>H5NA@ZU8ZPuKDbi}~>Pf-"
    "g&D7C*}$p2On27gr@)4|W7R)1viw$<k=6up-66hf7qB-U?}}N+P-"
    ")iz8E7Hvt8Q<_jQ0Y2I?JnurE%#S33<@{SN(#cD&WWazR}_>lS_gu-"
    "1Rd}6Xy$#4wc9}dOtlqc*iS5D<k%`G%VgBrYKSS#5)38akdAUYWk4guM!Nf;kMM!^-"
    "yNVBDi0{+{rq3;zZsvbIxSjf`jCBs_PmQ0~Jjeta#kZiue;*{_7spDBz+0o5omQ$wiMUVDtC_6Y-"
    "RWa`P8@U_(t~Z7By#`0j(&Z#ek9+a>tGTk3Gv@8bnmbI|4{%N~XHV7qdGnC-"
    ";I8|FsMdn7pP}sLP*s%RH8dTTTlQK_zwLUIc_JhdHZPcy&O6|eoWajS1y4m0n(xM}-JL?-"
    "S&xN;rA;%q2Fb?PQeP(4%L&oV8|;zeENCJU5_UGt+ZuFFoN{#B<B2zW-0M4@g{cQY!oC?sX@%X}UsTZ_-"
    "j2NFO>2x+6a@(zSJK|Y#)hS-+{EMikT4p!YOanrmg-"
    ")w;h%q0wG&oZ?uOrUT_mALLBhthl&*#7jN#}$t4Td@ErMOI>o+=6aV=t4C|EkpFz8RR&4I^kV>+j+!8<#i#;w6X!nUR{&hOxg#KG"
    "$j7Q}dWFn5fXoG-_TNGRCZmL%vFJjL6wsWrhmf@Q8;1R@D;Eam?H2sdc(XTO7+^`+_-"
    "23||Y@8s@dQNGaQCBaz1X4BsG$mE?vI2gWH`)=#J@0Jg5ydStkv+U(d#7crT+j<7W62RE;0IC&n$r(h@lHkjdxaJ;o44#MpdxWW5"
    "UhWPW-3b*Y35x7#4{VPp(|v%K$3(ec85BL8kCEtdjYUp^!HTt^9vK{UEK@yZ@6p>T;4a%zw@e?K!BWRfH;7|TqH-"
    "9HH3ctubQqc9k{Lb)GHrEGq|Tt=axclHT5PpZ?2i;FZNpFEe1&6uap~$r@RDH6-"
    "b9p7vZg?2sUB(6n|b7RnRT&*OU8jFL0;h>jwZ#>an0U%lS#^pu*@|fU58|=lS0JSpN+_k{UR&(Ucnni2OZSIZrujUtu;Dmx|qzDK"
    "!Ve<weA>E_7(i_G9Gd9Te-"
    "D<8%<k`3Y7#?70XPeEEdGY8O5qXd+LlC>b7Q9#80}Yt7dg{ZHhYoSx&7bSbttg*<^7Bz7_nCw>~DYu<G%#E7NJW7>6ik_)3!DQ@x"
    "Af0<TO##tw{hT!gI{L&Q_;#h4>*^2Wm7us=#<nVuyL7K3q+9mmzb2~dcc+$6WSBV(IKsmxy&{R7kCAQd)Zd=$+cO2^VQNmz_lnq9"
    "x`d>KkcLd@c+yE$IyCzA5mHrE51QYHM0u`@<-o7&Xa5gnO3X21Q{j{E!D0EU+7ZAG3lGlXRl)f`b-"
    "v{ozoUJfokjE}+aJDaTSqbZYBp%UY)WOF`x+T9<~JaqDILe2L0>jQYX*oIz+FcUbkr}5*=GHVAOhR1wlqIKggff&0L%YwTQSmFQt"
    "kC?_5Os}suePjZ>5yZsUvl*h!6k(WQoO&*fyA~tXS%aoCqi$~Eb=8#N?19*Ws2tdj3r;Mv-"
    "V}n^8XPfuSCi~L?#1J;=E_#i7_lB}?l4e2z&XV@HC6MQg3>}i`v9Ik=m%43=$2T=?>ZI<(qkZD-hz3MhbAi)73K~g!2;)aaEBfT2"
    "`d*<sZoeod`EE==8X*)8RuS!B7%T~WlL!-"
    "yOZpcuuiMV9=Hx@CL$qW=dxKP;iH_A7`=UOgLz#GBSb!&gheZvbSJ|8qC<V5F>Cu%)t-"
    "o=Az|*S*^N7paq4YN=jSq4sHrzu+qKiCK|#X6v9sisJU{Nn?(M-"
    "4)sfqZ*Tw!mHgYFel;IX%g+ksXFL36P325_Qe><LrL{wv@?2N-"
    "YO>S&;YF4WYmcyyZoqmsTwfX7rQh2;<$>$tM3!G@CuP+O_8aT+(KqbLc-kKo6pfw=a3jRw~Jc6c%+emcM6VQ_2%U)NQ&{7?AHwv4"
    "m-f%A`)LDEqoneWvNzi90AD=Evih?qtp*&k+I-^YjTd-6<hpqf5*2+*{?_=qm9S52Od1ag09%<t-"
    "i1a;X9JsOP`rB9A(VpD+scjR_tkl^b-X?zXs@62fcHtg=ucp**`JyDEgGPlc0+i-sR<kxCRU}wnLD20xFw7TmkxCYZ&!u6lriW-"
    "Y>Nw~Wo;^!~IvIkCSPpYUqmZ@~b&Tl@ESnm+yoW6Xoq+=Eu58UXj0y`SD#x`uT)L4T!ZAt6Ww;m;G~1g`EJ~QSMk4XYVJnXEXmVZ"
    "7Mk}}KiTXe^Y9EFRAVFu*^6_tMp#!dCAFYZAM<a{hf<Le?COE_V((hxX)hkV}hh?^VI{YLEE!&JwkMs?MyH4Oc9IGCVXnI%;P7_!"
    "Tg$DE_h_n=W*%`=ON=X}$!9G`7!9k#twKtkrQVrvPlOV8a^Gu(vrV`Ofo@nu<#U_VzIii`hHqguq)zZuIh;Ez2v_V%S4W^>Uqk$5"
    "$b!|-d$hg;!zosjjJG%tWSi4K{9SL>1+B<}KFmyJI>O_Xo@FR`3WxeUUwYeS!6(bkYSK0C_-"
    "H8|1)|;#`t<qWvfr@pD=I)Wt(sw_GtJ_oag!eqqDS~_1aRZ4S3Kd(Ivfatrglt$wJH@;C_q2?hefIJDxPYcV<-"
    "y1A2^rYb^{^+kM&1ItJvC9CTWR*)&g9h~pkmx=h9zPfuJj<jL~tWdYk?Sopkm@$DuK8U^6uG(5aeskaS?`y@te8VWDnf70-"
    "Hb)vAt;uvMgNvdbtAuis#Dl(9hjw&W|T504la!G`3lD4@%6>hUSUOU9}*6sn1%uJB^m(R_g>DR1BR@)o`6WSp@2Q+|QcU!=Pg1f_"
    "bHQ4zbVAa7W4Yk^g*4)+ilAWc0Hd<6HL|Z7~6|?zvZR$bgavHx+xsTrR*4$FgK%SJ_|tt=wHSoo9NuR5&YH%Upeu_wl(#yI4F}9W"
    ">c06)Y9HEX`#JEbtTj<tzC5Ngb2q<<_uNENY-qA<I%$7>aE90uwtmuLWCGA4#E`ijoRRmNLKLg@;{0{Ey$rULYlV9}u?ew{k1%Sd"
    "#Y!=v0WTS?qoZ<8Xu1^@(OzTV}#pyu&((8ex)=6XDU8fI)`S7OtKOipMfr8U|Tn5hz%r9<;D4X<!J@A(Ato7|F9GXK)7h?19XOY="
    "J2bW2E(Y^44#)8J=d{MA%dqEZXmrKhi`<c&_q3B558SS}K&4>^D`9P0*q1sV0}nvn6PdE@6$Ov6^UR+m>iS%JL0#3nIBX4_l%IgS"
    "lI}vo&wwM02jzAf>`p?H`ID)C$!s2@^zQUBx1EsgalMqho9YEEUEUtwpS|l+<F0T9K~3wWV3`PKs)54H|hy+Ah6LuI?jR9aX4Q$S"
    "PP%R=<!PZ~KEzVkKqm_dsaC>!Hb6>oHT|u4t`SPqcc!h8y@)KR0$GSAPPkNwaF$(_^H<m8F=+o3EEIWUOvW{xcfQU*VZhh9#Px2$"
    "%|SWor-F#MRC-d*tS$6uRa{#4)M+4dZ!@P@9~+0f||9LTB}e(QX*EI}$Nn!}=^|AyVz1&M%}Ne|VBe!A-"
    "7jb7MN90U98xrSl78)EQ39X4YQLmF=8;0OE;BEMN$9x)Ts-5|Gc(lo}xXn=a9Nqwl)$SC4{*g|nyx{0{#4GS~*<-qL-"
    "y4Q9zWXqdTZype|!Up~OYd!e5w#M*u5cXI+18kR1Z+f>Dqy03{uuE1>Vrt6V%8VodyTQ(c_0DOTnsiVB33O;ia2gl%Luhn!sn>G;"
    ")4O^S^0!*5)12I^aHcqGOiqpU;<|1yvRegY4=!;Yei!K>;Key@yW_G(|j0Ho(+%@yO+K&(oc?XyGS~)THdaUVfxSmt1A|PQ}lRN"
    "--P@lm!x;^iPKZPAngMfx<7c(pp&!i~*J|ug;SJ#P~ker^<BB238!s4Zv#f99hn8u(R!5NmUAFJn*8=*K~BX^_AIuDyMi$Bo(c;r"
    "JrN&jIBj40aAUk?U(i~<|e^k;xORZc9;cnL_d7b|=tCF1|2X$B0gp_ThJn_ZkxgANO0<-"
    "+xa!u5PPiLv}b_<LZiY;T^+mh9u%lM!z|&@H;ybsN20oi%(kJuQi#Ni2%xMQiMB!%vRkwW6){5dL1Qb9>rakATeIJn}&n9SxyruY"
    "(4^dHY3;FW?Wh;-"
    "R{mSSA%Dz@@n*mhBDGJ(S%F2b$`grQl+g8KPo24O$w{T+yQZM0Ca*{PPw6Cx62uMUz4ds$26IB4G*>ej1FHY&Pp23U@&rq1EVN>e"
    "C&9q`{Ew)-y2&?)2-$Uqyf+)`NDt&vJKsWL;<))HLX`R(vIAK!daP0nS-"
    "by}8?bes!H~cv$wi1jICWv(%F&cKA1<-zz5ef@D8aLDFEzK_^<+{;%(p5$#RPdxWzsO@o~Tj}=SZX+O$n=Y>9Kzt-"
    "fX#7#6!GkUx<C|k5PHc%`&N=X40j-mi*P?NV7W*!YSh@kE-s=EkF;z|H@cH!nklr-onSZ;k6SMd@o_a)W<f1}@Nd&pYgRm3!SE85"
    "HrWXOdgN)FIYn%a^nPFE;^MI?k|MM0%(v&|w!#mp>wEJ&8qRIEwvq^(at@<7eJK4D_Utj0H<g=0U$KMrH!YPjPjFJE1QP(GrtM<R"
    "9_Jl5<xB8JM%9C#v+@wp--i#)F)PJAFnn!~T2#dMm5FyfiKsIdhL0i*s#v&)|Hh-8;-"
    "vR3Y9VCpoqK8*b;na_j44E7{4C%Ft7JTKbaBBg=i`y;f^quL)wu+relL2Zfn`H80E4uU9ati9gH7*~z~r@>&s+Tfs6txs6S29fNM"
    "vZY}>l5tFGyo4O-eTFYZO0Yw-TC%*OrEBEGwX>1z__QTwoTbwlUA~my17}aj=VLfzLWJIV$~d-"
    "8w)&linDl(~kZBNCbyBJTL(G!;2}EkH2_->DnhH|jr$MK^R-"
    "X5R0gv!!pM@fHgvf9Bp69L4u%jmtBnh^zpnj~fHL6+?&JKxbr@)YtnNg)jNP{Qa0g`>$Vt=H|A$Z6f$p~5+gypSi=bl7^4e#JuE|"
    "iw4^)OD7yW2-"
    "2pqmJs27?6)vFELDOc`IDp2Hu+(gt@}qs@~B4?xaq`f=q%h%{I#+Ej!*lY2vOA=^s&yTP}w)=b49Rd_TQDmig6#>=yz@G*#Kw`Ec"
    "@R<w&hg0+rf{r0?Q<Y^l3d_(+YA{XO6FR6rX66`dXs#trZQma1A?1>^Q31v^NIJ*+sEOo&N%cNOUEl#PIJ2Rq)s<<wulq((-"
    "NW@Z;ciB@>n(tbuScMQX?_!d96JI&*Dz0qUjQYgM77oRULx@vUCZ=ZbEP}+(6DH5zbS)N`fC9U14N}A+w$4{ErHpT#ITm=1elm75"
    "nxuwL<$5@_MdO|jD|ahcHx>(aJC`yn1v@h75R2aOF(cm87J~>5>;^W84Zo370k`xm&MWJE)(VFKx-_sJ2X^aL%)Z06JivXB`}VmK"
    "#b(gR)eip8vCp7~g5Aajh2fFomJDqCK@{>wli%Wi7#tV*45^M=LFaLhu$wz}6x@=>zg?Z$@WpV$;Xk}ot9@4conF)L!W{G5lGpLO"
    "$(C>PUw=E^{)uSDG*Ia=PKgLx+GvaEHC+#77wxu2WMqlYE#!^A801_vNU70ctI_SPrqCl;t=1WuO5zg|3z(%gHvZcKImGrq#5*X}"
    "R1NNT;t^-npv6vA&6d<bI0rT6T)~vqC5A=cAPN&XJJkhi?eK%6Q_L;nhqq-ky-"
    "6{k*eSA9Af;}@;~eQF_5l0`$Z`gL^eTIB{no6hGhYC9CiC_mw(+4;bdMaF7^|&LvO@qb-"
    "1kE`F2+tzORUvo?cjMsQs0&c89QY~YuW~r#2T<@8&C~O*y}4BE;pvYCtoQ?gP-"
    "WuOOg>&TYY9yqC=cZlVt+b@`LsUPtHfQ4G_?p;cBq+Te1J7n({vAo=i#!B3q4UFCff%C@Xe%l?Sii-o&$tYOrEw$x@^wb--T4>rt"
    ">Z>j3_woYgGFPu_@`F7TJn;pcC{D}wu1DLrll{|K&vpUaUn>YDLmCvwqRnDRoVu6Dq4JC{=UpVhv>wI&xfs5hVRZe6c~CV*+U0PI"
    "X!in46Um+Wb%tNXm%2MJ{wq}VwsSljBN$;$9e4%AsRo#2SaC4129c8RdDGgoxd^E5t;iK6F;Oz5&^Kaj%z@@+hFf<`Aq0~I?-"
    "c48zW1-Ie29sc*v)oEN|<!_xmY+{+pH4Zp-?(8+G4RTtqWFY85<&M|#tTBF6>&<Q<OvH(uDtm<@N<+1-"
    "z&>ma37<ABH6w|@v2$0ux=ew&S8DpY1jJf-"
    "@+rI0MjI0nTLW3A*JKZrwmC;VYKeei_Rc5S+x!;Bi@UP*Gq$d!pdnnQMgdQ;k!?<BG4;1?!%7g7w5&JR`;NOTBA-mlsD{S2!StX|"
    "#qm&$Yx7}8OOCsQ^cZN_+0gkmG5~EPqI;~5uE5x~Wv<WUJ8gkW0g4t0rnb099gePI$mLzGemE`Dmdy>a17)Z)2%3UNp>%eFe$}ey"
    "zGXHM3oWY~8s+%ay$A1g4{ocdaOtWaEZc6|<8%OMS-6^}@jF~1-"
    "^nROTX1VdtzCx`>TL$2Z5zJ(p73N)v>2$F8x_yx@Dg!Oi>=<v43x3IX}fMu)B_=9@5P0*&9^(SbiVj{@jZk%F27#<zws8_deiGYY"
    "_e9rH=Q<IJt$HPRPxp$XxlP~C~nE0HAhp<(gWsgc0DxVvqaRiXe?M)ZB5etuiyE%<^SXUH~)%Gt<|9YkD9)slhH(CiBjeg1?;VTQ"
    "xCqP97;5VOp8EEN!eGhk%)rfs9Uk-w}Q)Q;~-"
    "6iNsFtp&EbltBaQQ7p2b#lxqRZWPP30{u%aF_EfOnM2R$qEox7+;JCozaUOhTgo${$7`)cbzfW6g2(>kDmN{h0Zwf#$!dhCoWF}f"
    "*xYpfNli#;?`_6)#Dk!h*xKQe`PTh<Dzql25RCn(1TM*!2JZ_(=327eZq;;-uLjbNL5(&(Vt=@hKAXv^Ca5ju^JC0*|`L^DZ;7-"
    "_Lou&h_}Le{44##Cz_;t^@k!==TXz5f67A$%-"
    "d|8o%Qw<Q3OD$R(c4}^3+(pICt*eyP5##siNn&)HM(mp58tVGRJ#ZjMX%2JB3XLVRnfwOCMeu^<{wfq~4knl+sk)z5^Xl3%EJrqa"
    "tIjiooB@8f{{en~`c<1qv&-Er25!pn<wD?=JuP}QC6(GBDh1u0Uq5zo&E-mKr_Mbd4I#e@h{YD?*Q#V9Qi@1W#(}aj-go@Vkh<q%"
    "M$Unp!)L{MG*x-xkGV5YV8fCZ=T1*!$&-"
    "{i4Vt!>G;xoVNjCVRA8~YsrOp7~vLvLj#&6p7BtOX$S5NQ!swno_K)d)A?5Wc?Bb$RBNy|y>Qmi2L9X%T2oZeV;<;>@U`+p%bI4s"
    "P1s#+k2$gc~0%&cMys?v7#^B-?IXNU1^aN;<|-36-@j>rFv$AoA%k^A?iKtAANC`YYQqqoP29Q9~60KsQDC04hc{=>w!`q-"
    "d`57{h(X4ejYWB=A_bkhU%COuBu$@ounH*7n>Fan|5KY}lfCR;#GP0(l+fhP~nE?z*4Rs(wCVYnC$fp^I6Euj3u)%a8qysdb<aQ6"
    "M&J+5Gem$~;Ber6Pyt=j!fTKs68$+q9CwrU~(yy6=8sdom8hwykCuw|Xq+_io8J?*$L?UL2_r>$?HSI_vg3Gnu$*fa#tLTeCe2Rh"
    "yj+p*do$3=CBp!w5>ePmPNV`fE)ogvjGzk%rczMa<HoH67cOo)8xGT2d~Pt)%#%Q0_Ze{J+4fj$)~C)($)Y4@;TtIM9eGE7<Mw2{"
    "Xv=V<VRQ+)8uZ_i*+3MA0H<sc3EeJA<YF70lmwuKEzGAlS5qrjt&O7coyI%Q7}80!z3QY=}glED7U>N@IKQX8~`!s_eq;0QQ|m3y"
    "s&O0gITV%JkHPhCy&*dO~>Es@iI49M0W7hNWaHS~0pms7s4wu2-"
    "5}uay(ku2`tdc0GvswU^Z%mq49a3Y$jCub47N{agx3ya@Y?LbDvlC)u)x&ESD~ycxI@?z7hhTrs;BlkA@Mjm%eiWkxb)x>Hj#RL~"
    "#do?;Gs9;wP36OUo(Y(f`|sP*Lq+-}p4Ius-gUPx>3(eR0w>iGqJ5)1nP1<Np+)N`!loE98OW{PQ<c_V}QyKx>?Sj%@jXhM&Jq{&"
    "O>vWLNdxovP5e+0Mcd(BWh*LQw05|XAjB=h+d9>{)dk!jq~bR(}cS%-^Xvab8t2;&fuq`)NHKS=hTg^{Go@6{~YUB-"
    "pAH<{BBAjv>Tn!ReCmmr$G4gu{)$gYcc-YX?I+<}Y>pCio&JUX7CFb)(+B5LL`tv7O3<&j2IW<?DAyU&~tQqaJlXmOLQ|1dN&kX_"
    "u8vWu@{u_cg-+?dMHrvyACIath4p)4NoUu1IbKo_E|H~Fb}YNg5=NF-"
    "^<XDALEc3%i(ZB1F&4c>WiH*ke!Q5BYmZ<fs1suU|O7BWz^C&$4&m%WD~qaAG#y5;j3=U#iGx$ez`v2ifxg0zb5UULyfJ%z>S2YC"
    "Sa2&Vsw=0Ny3cdyTawJ~tff};ttHgxl~P+U@N&02Un@<^O3r%nFIGBhd?5tt+?vy}Be4f7rR&yFCG9(7gj@dxrg7Exn8T9W)(DhQ"
    "C<k=j3T9@{l}4+oM5a2#loWLB+};E3Jj3_Ou-xRz+B`Zm{3V(l&(ng2x0B#Eq9D^5%s|G-k5I46M;OOqP-uP)@d;7nfKp1Q8CH9?"
    "4uJ;GACJq|WWIv1@;Tu&Am#v*$?Zy{uhvmGW#c40Y&%vDy*(jbUo0roptR``3pCNGqZ-"
    "}ZvD2^LLW19XzC7M+YfI)K3EL9EL>cu@Ntl(Ub9h)L2|veq1xp{@f<htymZfK=^G+uWC>+`LFY@oXC#0fj8byXLVxjw)P|>{Z-"
    "^*Y#m^l+kml%vCKjm!cXG7k}i-7Hvsayn~k<!XJDQmKzDt`7p8tO?Xa4N|CtQg5{1Z2#)YlrO-8q)yS-"
    "}>#T#U?MefhB$<nr$=sF}IgHU*M|0fQqb13orP$5S@`vXTf!>R^@bl&rh?JvXP>+`+g9Xd3%%m+Zxcr+r?2WNLIbvJJqPKM2WUPa"
    "(WG4ccB!wlrr2-"
    "iW>`+w;V5_*Bn9s7C(UQ9PktnA|Qsu2;ZCtX@;uXpXgM#c`eMWMbwVJ^*xhy)U@gyZvT?R>>ZO3_x4cNFgO!A%D7UexA)1H`_@W="
    "e=9F$UuguIb?foBlsI!K%`veFH3dY#-"
    "2W#EP*kq`r}7$c=5M$*2K`AV<MNJf#6)YOC`ArahD6bYF}s!SpwTYzmuyovV*C(u)852nCM3P>cGFsX+ego<{G(U-"
    "YAi&J|?KA5Cwi{`fQW7yBOK(Oq&8k1UOxjJk3uDifB7-"
    "ng1TMc6+qZWwl+K9y05}QqUaF92gxMahCXg+8ur<Z6Ex2JYS1ux}UWzL0hlA(}Xl1u_2Z^GqZI4bT{Lggw)$T~guA}t0;Bxx`ShW"
    "sphuMa~>*^vUV=oaDSmKdmV`fk-i42VdwP)pAOzuI@O2K-TDe}8j&){i2gXm_J#$a~qn{Brp8%g95-"
    "`rmX}=PBFhvYu<P90R1e*qDtvBCkd?c0Y)QEM+_nR|jTK-"
    "6}WgQOpMe`BWb3ALQXE40S8MYOlxBow+bR@7`V2z?ly=3h8xqr0}~=PUp7UxvB{;=7WM_2BDV0;b|c4v?HaRG{?EM;PLR-"
    "Y-X+*kj@8nr3}=Kw1<xA-Vj}_ynkaBq>aNk7mQiz0N}>i0s9`xvh4S2RZhL>w*tqx<>M&INrp_(zqZs5Ff_^=xT4o_23NcdWc&@G"
    "l4P!Ct$=lc@Bx-"
    "u)_IFF*y<730`MX^CT2jrhIz!tT=GlVD(hXowEFD_NSantfG0`qqP3E_Ny4sJs_5n|K5MIyAY&#2TUYX2q&#f36pYP-"
    "V<~*3oiMl66CiWi{GT5+N%b388jv+WN%CfIB<D!olhC<<-"
    "ZWh`Rvn8***MfB*(}@N{KV2EuJIQUV~Os6!7{oM;KjEEYLJs8)l!eZXdWzEb`AxzX83av$+OidpecFOR9<2^UuvL|d{pgq2~1XeK"
    "C>=?`73lTr|c4tJxk$sbEL-0Wv>-"
    "`%A>B%n*;Q4V)id5+28Pb^L1U>>=|1DO!g0b&IsHob^y@lBX6!d6s?%L*54HtfZ=Vwj>`;H1e6S1$WT;!py~L_Ewiar(jFunju4x"
    "-Xf9dXl^_#Fi2PgUajOZVLmU^X>46Zkwk=_SOr&5rpU`K|QT#x+VxbHd5TmPNUq@GXnEY~79o>7(!zZE}flG<Hs%5>h7y1Mh*DJf"
    "(M3bS?!==QVy%T<>gX7E&_SDZWm6XT$v<;o^Pw*+*Ix9&QH@trAC0*Ir84vF!tA~g8DRk4OLq>#!{9IMP<Fs*Mo;q%Dnz%UkP2f1"
    "~UYgfdCY<Ij&wo!PPQzEs&Wael^IR6e9&=Al&7|^?LBScT<#z7+STsZoZs?sK_5z$if$?s<ZDF<VRyINhfrxQyX?IIg0XpoKVFCJ"
    "7Hp`Sn<tWW=tBiaic2EsarcGns@}RlYTL!Knb7c1xTead7__m0iLb6Wq73#;;$jxM@!kGnF=i<m{EA5*;XoBLQqa0namFpduV8Xw"
    "Cf@u1w`mDj+l{>8csE=g=(m32i7_8XMA%=<7zi4M3vCRgup+`};L@29T+vX-O=bzbDH$e$MP8RJYfzzfPXT}-"
    "bJXm2NYXx^W1r|UH8H00OM*c4AJ#4Z~ug#`e&#eci5>Z$%CgKfml=5<~%*~`_-X@Zw_1+`kXHoP$Z=|J2EFUL}y>npBwca%l=-"
    "Rtren`&3X&U^SSozng50efZYb__hplkM``8JQ`CLGII?@+gCDSej%|KU404Gbng9mYzC?*!|7ds=-"
    "D35e*jVaho_lSS+yUUe{0tHrvGtAjKU=-"
    "Rtt&V0TA$S>fkj}&|REuJp$TxUL22z2djQfc3lg~$hx`M8kH$I&{tTfsV0Z}vO6HP;&L<M7aBp=Q48q4DB!2QD^x*Dw3r?(T1T6#"
    "-duo5a}%7ppC=ScXiN`|1i60bO&OG}lE&^$=hV0wAKYMD-rIARKG4#wJWAce&6QK+$C*Z}cxgr|IkE@4%IN87#O_pR&(#_uZ~G2^"
    "i?wypVPz&Az6kY=wuTnKzvQH2aB7VsxxsG+SFZ%s;|GT<1yhcKkIstKOV~>Gf#n5@0&QzLmqo<M(<qz~a{4>W9-"
    "cv*~R5PNWf_N@5H0Ehp-Zdb=N3J8&M?o$sfKVCdT4G-"
    "&JvoKFtnmOE7~f3w+d=T@41*W=wZAn01WYMwTEDHRG2gj7~E%**}uW#4ggiyj4CtJl){@B1=KEd)y6$DMb+)HHId{&ZY-"
    "4IZ*A<QL5M{uhA!S(hI7T20USlvVZf=^A@6t(MaZ8Rp)N`Hwf9=Q{K#=vqB*bZ9takvbNM%a_uV+frjRk8?Nr9X@}>u~#z|5nU=w"
    "3DrTBipXRh$R+nclU%K{Fu59R0L!jZtY%<zNhzic6K_GiPCRaT%e9-"
    "YgFx5b6AjnimSfIN^+ojcrpNs0H&;hs(WRtpF7()kv(EWANwl?QW5c;l-"
    "B>Ji8K`8dM?2DN2w~~j{V9~&=%LWHe$^aW8w7F(N_(G_7+aTx!Ogns!RCo@=rUkB4T~O$!O>H=9=aIw9dRCRH#!ep$CwS6$dXZ5N"
    "SVwE`&W))nXEfg%4CfIplj{LR3V5xwbFRid&H&?de|dgG+fA=ch=`{BYl<u-"
    "9cN(mh|q9p)=^ST86gnH`>q<n5%nk%h58Tq{~(z{p{aC+jvq=9s1mPQhf$Pe-psXoES@!LIS3;4A>rA%wW=AuxtJ;L2{A{h6Y=Bz"
    "tNh>wB~VG=ZL=21oP4PPu{}C9LG=LE4;aHmNU_#=;ny3at5kK62ja|(&VLY@FU7Bc+x<dBa$jECfOO!`3UFO%|k$-"
    "U`Zc)G?6g~8wL)+Sv;?@3`Lko27-"
    ">mO_x+lGtx(k<v~Lh`gAmv`|C(_LJ|R!rLt&eFnZ{lAF2n3B`mWBBKYXCwP=4e9Yla$#g4&yx}-"
    "JBt+5tAq@!D34Oc>!b!(&0xqVLHKaqgjQWH-"
    "%SZ5v481w|ZbeSwze|~wx)3$uob2XQ;9A;5GUfY1}fz?5$OKs7<h|>|jApXyzn3NNH>h%%oIH`!~(rHNuZ+H>#RU{Zsl=MRYeaYu"
    "j(6ES5CZeXxrzMp<k*N%ewG!2y?#N0qP(ngM={VqYS*_SQh=#HCgsEw_f-Ar>HZ1`xT@LL<hz_B#;Y`&lTIz4&i3UY5(k0DK1c;Y"
    "ozJ<=h;P1bO^V*3fPP@|db9FT5lChx4GFn=&_85vo_^`P}&78<C8zfJp1}<F|El)8B<uo)&N;DC64OF`HS&H<CcKM6MdXANDAIdp"
    "U5U2-BmqRx(o}b`VC-SWh<6)i^UZr`*#~J6*z^BV_@gFmP=fL5y8sX6wl_Cn}P_I5Ml^v$Obk)~`bk`k>wFL82w_U0@r1n(PXDE2"
    "R=VMt}#{s9ys-?Kj1RX;tNS(jdqEgN7q#6Ea&!jK$R!9UWLldBe+23gTxMs~HAScUgdBIvZFHM6t=W-"
    "Q66mM=i!ONDyKEqI^hCdO3y@h7i)PSW+VBY2=8l%q6ET_=SjR;S<U@fJ54tMat^%DY7YD&2WUMsX>VaX^b;-"
    "*WgrPvNs2S>wSN=M!LAZCYxO+8?`JX(tGOrjl!g3q}WnJbsA#q3h*Fk*?|&68kbxnj**y?_nq8Q3J+0$cC*g0rL7LgTBdQ0dZVDX"
    "^1Teij9T^Q2Q-"
    "u2~DXeuj(N6Ru#i(8J@$eLQU<DpIoSRjgI&WR%}V!sNW^bXOLw1<J>()?*2a&xgEy#ahTljo^e67j&)aF!y0b!JA~5bUCyX!hwcE"
    "e&l+sbm;`xH*h?65r}jNE7~lOGYqsrR2?@TlW!^|`-"
    "%QR0q>tgV}3l>&nu39%{ur}{Tk?WsV!TMw`HBRhAZecXG7U4mLfKXu#)YFLPA9-"
    "qSZuFFQ<Z}%Uac%cO!E>{AS)!e4O>*kRLpa4K#j^3YIQ`HOplgW5=^%>8+ZZIj&WERW&;hLO#=Xr`Gvc=~eaw5r-PQbZK-Tx-"
    "*!Du94WOcQK0VL=e;EGH+kGEbZ8FbP8s3GoQR#u+D0z9(5`Oqz>Z5KJl;|Dh;^l@>;aEfxdt%(~UW}$#>72m>eJ#D_#C9#e1H~;W"
    "Xi8sp|2q7DKaVYOvDfuWU<rE-aN7fB6o6_=}tgc&IAS<?i$wSh_gO(CM=4Mre5_uCfeDNKGH9%JL5J;9Rw8jmN<_uoKNjDTb1vX@"
    "^X}OP9u)H5$(#$1ZvSVqN8{emLzC&(o;EOP0o(<yI<wHHYxSOSr@zRLNEzzgzX&5BeQk5xfN8bcwYS29h#0V5s1LwtF3Z^g_(O6f"
    "aPkAibx2>%0z!lqX2eoGCAZF57kjMVB!A`r?CnV!@Ft89_&vsDiaK=3oxkHoU_-O=G~#rr!;&rjg^-"
    "tys>RiMZ*KTC_&07|Z@%WifZ#N5kI`B3;5tmQ!6KfO4uv)FIVX@Ai>IdlK-{rLk<^dA-"
    "!>06>EkI(PMw`q`smt&3dt^d>0?@>XQ7AULy5ay>x0)K#rf_Xd6ru3Z?WRMd1C&v)=6#mvalL!?XI6+$<IOEFy2xJ$Vi*f$N24`a"
    "VdUO7WJ#@U_*28dC)xDX59Jd?v`-bt-lA$f2CkiWp!kAU|mHchAB^#YvjQGiiiHnhJO^U-"
    "6L!$d$aG8U7_NcAq`ukXqPESjfGD;re`Cc^fa0bs`irqcqX#f)M8H-"
    "C;V;~T)J^WM*eM7W%`u7^R()CGg7w`5nynOGvCcJ8uG*GAR=pk?8r!NPJ<3I`awlVP^YUEi^dH4tdoxMZ-gQeXH*J!FV-NYAN}^%"
    "!W`xooiW2|TIqWZmDsbheHZg(*HRp6mL%wD&kHw5(n+SY3ARg+ZbNefil+SGcoKf#%Zi<KWOTeAQt1k(!;PI3lky5@AG_CQk%H%i"
    "cAEz25@HeMkWQsA<((Y4&@(p?1^thN%Za%Giqw=C%A2Jj|i+x71SjIW~Z9Wj$C?mK@*W<KWOTyy4u*@1|E@FaD+qGpw>+ztM7P<H"
    "#q|GHTwuM}E)=dh>43w(ITiL$V|{YNHm?99^PQ%9PVfj!#r`esqaJpk?M_hSSTtuzW3EoAo$%+x51W2!xisOJ+AilzG0E-"
    "=8PA)XjCau@QW6?^GKw(ITL1yx|AZTiQ+xPz$^3*x}HFpk-*o$uW|XzJQAr`H>$qd30V7=rWo;4ht=-"
    "n~sey<b>qUYPDXw<8~E^d@?Pg*3zHF!%Y1l&xH<aMc+c2>mZgo3vTPBU^(F{_*1sQ`)8+gQcnSs6h--kw5!ctD3X~xy~C3eJf^bZ"
    "Xz<Xo{9=Yc)v$**_8i=DyTm8Mpk-"
    "~tSrgP?<b4&>l4^9sWD}P+whDljc?;&e{B!VtKTDb9L0$5e>_$I^NJeabUH2PY?77Qwb<eH)m=Pr{hGGZSEm?WAn~<D1mn~QOv9g"
    "6%@WM}Ou40q*WDb8VfWPE*{BAN~ZT{<T$J3Swa(p5q74(_#_gl8UsUz@OzEmn*Usw$OyQU-ZT7F?MAO0Q)E$3~v&x>{ELhZNa17>"
    "r#`fW_z^AV_|p22d#x;D4A4+W0eXnIieXlW5wv^H7f6S%O)pzJ(6ODIW$c9(k_OITE9$HAt>W6AQGxexyFci<|#5CyTy;<?5(o>!"
    "WU9<n7=gPazX_9s4HZY2!_zK3bTfetS%>g-"
    "Q+=AoEFc#~ZrvK!}+ucph7u_(tFu+!qRYF!9uq##C)5_k#!oy%JHKO}?&McvK{>pW#ou&6#wcww||f|~uCz!3F9$1UCPn60Tnf0Q"
    "@!u_6w{p7=-"
    "zi$J7Bk|n9xqp%o@`wvI(4u_G&N9iczjhy#cqaAobyil)$M2J1rRnX$sQbgxH9I{3HE1DGkSFY@YgFY1`Ev5?AI^l&suqDSq#Ab5"
    "WyIj+NHGDJ;@p{CxgQ>PA2mL(K==@B225)5>^FUJhXL0>X)CbD^{f(v%ve27If<KuriPn=`vNi?d9HJ|`%KVG4nI6mZiy=r_G?lG"
    "QK|BEX$NZY$%&im$vX0-*-Nhn6k^q<%cNJ@5Y32+)7U49yseDUwfhH;(!ln5`y~3zsKO}S2Dm`?n-E&sXS9PE-"
    "o2JGXomdy?Vq|+Wg;JK}qpK4jTDxy+fY7APENFbzxoLZ;7Mu*F9p4UyaTYDnzA;Nfg3@+PqF`<g+AD3%1{&W%jAev3(N<$y1AxZ&"
    "&SR;7BidwdX^zh%$`UA;n;QuVOM`s)uRh44_ypCVwP0~<bRJ@vdY25D7J)TuhlBKjI#><_)RlKE4erUNK83h=jU0cA)~DS=OE!3J"
    "hw-"
    ")}B@4J7u(|BDnm)fMXp_Q{5KMbDwAjvDJ3eF;>BDlSm=85STl0IeYFuJ*UFSm#)7HSAw2H9k;%7vqV=c|?NvB7Rf;8&F)irxvQJ*"
    "7IjW6LY<_`RyqA4V=le^Ps<$lfZJa*%d(_+(7>?VWsJ_-"
    ")!K%?K{#GaH|^D`UQLtTje3iEj;8#gkzh*^%unv1C;`IG~B&Lq;$Lr+>EWi>$8nZhWZL?|)4=acMiev2`eb!CfZoWPU9^n<O3L!I"
    "L69r9G<&6<T5Qpo6J^m$ce3i|C09>0JG)UlFbA|5I>H_0pS%O22IFfsB_O|0bB`fjDGH2|pC)}*N{T26+-"
    "T<X}~P$tZ!bu|#EShtjFB1u9BJB(VpA3-uc=#12~W%EF;bE&uT0x&8EFt)+_gc|P0a%o^tF|ui*j&MO=0knjz?k?-"
    "~Scg4uz0oB?p<?T*dECW;y1A`F>y+bB7YDR<G?M`G&v1AW%36OG`?acoxysXqUAED5>^CGrA!2LeY@9QBNZVFjhv3eA+-"
    "fq{_ip6VsaSO})x@DaIQJf?wP26>j>U=9Pp4wly!l=oQQYI`_{ObO2@K8@8?NtMJr*j)E*OQj!(s8lE!tB@*+1m!eQqVTV*sewwr"
    "C#9BKnv2WQh2qTIU<~TtoZrUDk3vkwp)OioHu_he8Boe~{1XPAt;t<(663_1K|AAXF?}Hpe-)Vc!uaR)KMBoU`71!u;H-"
    "*I`o#bf&^ag^Sc9aYQcaSeAPKpy{#VH^1AtPbN4-"
    "tZu8lBh~hbq|}w_K5VPYBdgzuP@bBegKDd@BXSrTxu&<n%Tks@c|;2?iET8o2*#(q4l4B4tUa5C_yv}ykz%03R)a^zh#nS+H`1X$"
    "Tg@HGQEe<w7On8Ntu~B|nH>~lQBMKc>bywZB!#d&Dr-v=QYMSms`x+3u@mprV1J!gaDCJ0pef^5p;95s-"
    "h`Sor7~D%(h#kwvDI#os%Q*JD?ZigGF#0QlN42=pl%MTFKd>xNtHt2eU&{q!?KBkx&*vbNVC^pk-Cn5W}OwpYsw2L8!86by|a<|x"
    "PMX8jH@hHZ}g_nRFMcLX7xgn)y;2VxU?%<K4W7=3JPRn#Q^RUn=9sm5>uJZk!*b%?5ziKX**J8=mb1)pCGsFyDdF6K%!**g1M^tV"
    "^kU8Sl@6L98P;p)@5DC^{ECxlx$zjW_zLc%0VXCe%*0pj{y)R+nZ{9#LSg#dBH!%>rDEsrsKj(JqAivHr4m+$autS5PnoO)_4ucj"
    "lScK4IKzdHm;;?ISR|e0O6%7>$1}2yJx*Wo$ezI6iOzqrdm7>q^{6Qb>-"
    ";fZck<Dh#;V3;F`G&>6~vnT=8cI0G}wCDn8kZtuyB>NE#S~3~lV^cps`t4x?Iw8j&-"
    "NO51OGG9<Us^fnz^KNU7gWGtF%=>H}qFfPNZjmci{-"
    "SS4i<GPk!4}_AnP3O+hQ5pMk)PnftAU@#(D)rV}dxQo|l+0g9TW$<OMVg^Fu#Zlq!gz>)l7Wk9yBJ2RC_bp+#jxgMI&BOp5=!Pa9"
    "VkVF=##{D0#3toc#>7Msdlcu?lH%`BL+N_Om9*KI2{Rce^yuTd)&~s-"
    "S&#;5KuC(>4+)n8ATHeT%_WatTg+)CCBXztp`EL%++kkiD!{$oL=wgrelvJ1tv-"
    "i)XWS74%bV#&g5|H4VdfTI2bz?3n8257tGzQM{vf|EGz73-*qc%<g+Q+b<yl#4`)|J0VhwRg!x-"
    "uIFQn@^PLtQB~tQcr~I>=7Z6SjP(1Q=mb>FvBuWp0l9^+N%Pm;}v#Sx48=j*-"
    "!)|saGrX}H!Hr(iYX^Xz20r?d*YUf_Jhb_*za7s;BC0V%dd6X$0Y$%M>zg``uH{Rm()ESK;J<4+uCC=57W3in1!u(!oJshaqOF*E"
    "?`qe}4QkDJQ{^m(GXo~gT`4_xVc-"
    "1upb{BY^UpF<ajXuV_o0aTiN_t;88DXuE@kFyO+glgu4ixr+f^dcfwzD~^jw8WXiu{>_E@Rmc(3kqP)pLq(vGJ`ONq9E<?$YaLd^"
    "qp!bk2N>mZx(5dlkyHER=%MUXW<9e0_v@r=h+VNznsPA`{w-"
    "c)DQa3>%ZeeiSqsawS|(j)aXP@=bNZJy`YAQF}lo>ZevwzP4wB_**4JCM&_+ET}vv}OXyRI7O@uCS$sld0ec%K?vkWs)tGoGl1%Y"
    "{9;wO5wLRnyi6nAYcS2C9W2oG|TR&jD-"
    "QrdK1gM@dU(_sIwGMNmuNR;z6Tb6>du{Cu2hdqG8kIJ0~p3<fQK#La|iRZZNW>mZNSMrUtvgjwkx;Fz$Ke`}3|Kt)tpvqF&XRDM-"
    "vuBdaaNor(IqF{X$JH-uYK-kDS}0R@}$BduUjk|&cA%*+<mIZz`hrj+Ipt%n2Q)N{C9e}r!iRGq8eJn%a_V*Rkmf`zl=O+l*13@="
    "7WF^Q54@0ISFuS`)!4WBGrg-&M-{uC?d=bO8H%BD2MDwn^5+x}2Jq^}N2`<5Ix8>a$BnVA&*nA>p0-@#J!1-"
    "5FXbO@r8H=7=7yRHq_L!oT{VyaO{zkts$xb+UyU?V&W%gmd?<Rc9f%Jwg%vj3rk`=5&C8|~ms)SH6YBM2xP+|;%CLV^>=fOGY3=d"
    "Er1w&R9-JqQG5w^adLH0cNNZ}tGTxRyP2)xQ;Z;J<m~R`o3b1m)YmYU~&1;GW4c$`etUpwtlZ{?7HjW7j7U31#!wQnu@J*Wwp-"
    ">ESK<_xp{`R2uY05D+%FQE5n&=Kt{Z;{TJnyjx143B(<}mTP4j4GPM3H;D?pl2xKxGIshF?zaH<PpYieT9>sL*LLdjs0){Qa><1m"
    "8D(DbX7Nw4(QU&za1N{P8~#K7PsQ=?_j0^dYQwSalNl3bMhdCMET6~bDzhm}Q$C+FuhXbbbO@)*Q>jRQE}LCHtD`cP`cQw*TPZxZ"
    "GujQ1DBt=evx3g~h+O^)@Jo(0gG=wh*yX@0ZFkv*Ye}7PkSMcJHj6TT8oK3S6iiqmT7C9SyTf{%jjr$5Sv5eKtEDT&f%#0@LR3Z$"
    "XPiZ&*Nu)CD%ufa*&H9B$AOwFf~rPkM=^q$dzc?!hdtNyC*;=HhI>871SE6CN6qY5eU=sbC$Ou29NA=bx&SLp??K?}E(MRzCUZ?v"
    "T6l!F9HR)+D_OxI^mi0t6sXT@yFLu604Uo#n>Z7%Al;8Y1AgH4dQ((k8i7IC_NJ3}WYoPfr?}nq9m^Q1pHA7*S!9?_WKuoiN=D3o"
    "r{Bm4mv$BsK8XP5X|K+k&gFYDIdRCp=S^FC>Qiznz`Nx4a<`b{$sNgHC_gkM^YZk0<ly_Gx-YDH!4dpX)9-"
    "d%uTvTrlnrk>cEhBBaK#uDDkdaF*P4AX6wY;KGzBWkoTM0TGZ+9dKzu5Vp~Pw#pH2}%ed6<@gX?KG25fV+bJWr#wD~`eM9G$#%D("
    "S0_i60t=M#Q13Ja;~2f|~MYU?%DnL^3H@WUw^Iz_oLR7HrH-"
    "C>~Z7kP@?g49?vm3@Z=C%1ls<?6n3iRe_&DASYTG<PcJsqe!2`(;GP5!PaLx-omgQ%d)|F6<Jny{J^^C^J+@TWYf{U72{bjq71A2"
    "_OiYy=+f-@In^f9f~OnvD$;1O~1X-"
    "?4#)i)BvW;q$M%JoiJ_@&V32&^m{dZTj!F5{CQhtjSU=hCZeXyXvvn`R+v7D)ur(K)iT4Ps})5_nYgk|waEi`RZM3d3hS01BM}GH"
    "VWrGq#fh55pgk;Du6A@Sg{vN~(DP90Sg7DB(`!x0kY9j!)lb4y>HxyWPzQ@<iPgZR%wx@(IG3o?dm~?d7<YJ9ut-_x(GsSyvS4ks"
    "-G^-`A+)>J4BXXwBBYCDh1CG2%;cgq@wI(T2}u~6$D*~iKMzGFmAtic<9j%Gox~fC{dLyppgAZ;U{Yo-"
    "W$Y#gW_tCj(PB=BDIUxIjNMR>np;%7`b5^3?2S_fT84_$&A@Z*`JL-LkAH_Z_4`?sg8W{y(bb3L&SV*dklxytiZ8EaN5AlBei;Pg"
    "64Ia3Fbw5;sc1P4_(le(_u<cuV$r~rF24#7Su{`sn=-2<yZ&*gWAaiJSONd|J{B_dXerZZDW0??tNTTD?6#86f-"
    "Sm*!!y=ErA%PO)}SDBaRic4LGx5aQLz**dNHDf1HrY+T{Y={Q_P-"
    "h;CKv^5mRQf=3*Aq*y4MHbD!p;O22Ap2u7;L6%<e8e08|k8Z7#Lq&j9empa<z3h9+~BpYNRV9I3X?T?4rhQpYcceWd!4HX<ba`(~"
    "9jU5M^GOw0ugeDYgKtb*Nsyb9{M}9IY(cp^z%%fzDrQxSn)PTa`J<YFJl&vYK={436O>3VcF}7;O@>uDy;)J+HOV5Rhp{+To*~C5"
    "Z1peks1+;2yL`Qm++*yt1m_G?!TXRxh%lc$pG42gsNR6cyR|l0cfr~a*)3d=rPF+V=#i<j|DjI=GnZUfY^1c+>`dAA4=17F!(n!^"
    ";!zmCkIbTyOEX`L9zJ|e)0}xhu7Ji>*2+R%MS|0pwbQlbhy))xAQN9l>1%|FPjtK?p^FFt!R<@OAlNS~rUWX*FHUf|`VHIm*AuZ@"
    ";&TLrZ-+`#+6~?)$Oi!r4Ju`OaSDgC_Q(ZaY9k`4ujSsyU7&{HA4`aVdUVcMle~<G-ODRXb9)YKLrV-q`&o=!YPlUN2?5cu@nZ1-"
    "`_QY3?zm6*#H{-C@WGjb`<pAUqN3znhdfqH$D4x9|xYA42?c=45OO6$H^ce8lwvcvk&kIe(p-"
    "{_ltd*@nfZwbp;n^Me&(WaG0rV3HA+ytX<BEtm<?5{AJ2iMJZ1^cC8O1iksL3mEb7D4MVYxNWcdV(YgMi<(W#hP`@XYf9+@KHe%|"
    "|sSycC?-"
    "xO3luub$&+N6Cor+q+_1O+u$#I9$7GTXXInbFBQQL4e<^RkJsD0*Ag`xM@UNZ?szbj?cS^m;F|5t;t+_ae6TL?QGD}4O={p#?5|O"
    ";|s(TMn5HEL2u^xOwuv)30{LoCE^+2x(*8}qBr;r|B=^u)Ci8MGZ|B0z&SNiYza%3Y*h)!;%@Gh9jTR%tI+0*r8!Rn75&3xu@)fP"
    "7J?<2OmwfYR05FFqe7w-L(lcHmD&H3%!0pzrR7`%8;96`&4P6hN3Bf)R{T_zY)NIMI&LG;%p!R%x0OA=YH$jV(lYrdR;^-"
    "R_gu_yc?-"
    "M6A2e#qgQ)IGll6KKpFt;uuD=HSbXw}TNSpCM!5h)kNX=S1MVf5u%!(VRr31?;Gu&Z(SBy*mLwr^gY?$|yZNNI`JkzZQ6Ek%=$<&"
    "6|o3850CeN7nHrYPRevg2iV*Z;t7p83ZNM&&7{+x=R#0s11%_q!lf@>TWIu<XOlezDu2;3=bcE@lm|6Yw61RFA+#^gGCn+S%E$&2"
    "O%CvXS<B0J^J0HUbBJmyW-s_ofkqnBIt+YkEAYi-"
    "Q9=+I!wp^qx{%WgjH6kV7wz`GCKkI^+?qGNnhmR<NCui!9r4Bwxr9!{^<=e=ew&8<g4$I2Dsb{hT@?klb<_zYAWj^PH2t@1ZmyAv"
    "!m%gp&`^;E#<@KBv#sSzS6p$$&;3(j5kq%rmNCI*ZS4>j}7D`r8Qg@@paSMlK$tE|`bdaid`4GuDPUNlX-"
    "1z4feEV}rVczLrQIPjih2VReYj+GZP9F`!>86VimPfBY|=C<x683-"
    "L)8xE$!$op>CIReo%#ZTwnVZGUPdRHeSp=0bUXVO^XTzC|n-"
    "a#~}{SB>oi{{l?4EorEb#XK!aiT2Kto6u)e9DI7Mh_DxI@T|leUJlL6?O&<(9hKP)-"
    "o>?TxHigkDZrg>Y>oFw&|Std%*WPajM2D+kl}OE^lo;20AvbWT?#)5PyPIbAOP$iFke+O&0hhAi+O_qD(!&sgq%=CbPU)%~UxX32"
    "E;Qste+}A^pazie{4`(q*X>^W}wOu`MfgMkgK|$OU=Osr^z74{+h;Za~wG;LF2ijF|_TA7>f%H3T<C<kbwBFkktUd_}ZuAeSntue"
    "e43pxN5_vvrTJf5F)o19|E26DG2dmdI0?5q|-&IsRORB*8tA6!rs8xgRo{izW_cO!P$8N0|6xTH-bCSuS(FLi&**^C{97(`Ek(7f"
    "JXYC}r4ncOb6v9h^9W3pvhRHJegz`mNkbv(b|`MauNb)HiQSS5EjoDW_d44rDE$m?|Si+;T2*nR5;I`G!$nXDxYqq)I+ach$J^<r"
    "2OpDie5s+zPLR*OZ$$TtDy(b19<tGjkRvyHvvWLpAOFurC`h4#D6%yn7}rvpg3Ak+bI(cz4|r`)%N0+#8(wa%mC{)4f0JT{__#rI"
    "z`qzXQmB$Yo1}@GwBV*&p5x#>B{s!U9<?k1p?)1_(z6oP|4Yd?wp{hH)O-"
    "5H6YEwymkb86&IMWSD0kYMq$Cs<lb8a1P>Y(Fb^WtRWW;bgS$-"
    "76JT3$b=Wsn)O9gC^uX_*DazAbZh7cT2w3+uP@}UFP0NoTnv8?EZUkFoURT8OUEwRf)Pstf-"
    "}}gp;$uc7gO092z(&r9R}%ueHhL8Ox4X?W+Jb#*2FTd+W?&ol?D60sIe{-ohM4OfYjE6-"
    "}IH!D2O9oVr^?`uX2(K^C)IEZ@~ConBI$u)@;{9hwid9VYJk3Bv>@jDEFJ#8rFNIHY_A3#d+%2vNgl^N^Oo<riso`KV-"
    "G$a$D8!posTPO4Qk!<g2YK3j>z=MOH*tu`aSB*y(U-Nv=12uRaR;h&T0HQtzEu-U9C<NKgl(R&cTtzDW>8T0we%d_zw-"
    "$&$+NmBfCqoE_&#708w%aAw&obXh90`dYS>g0pJPVO&^giqhE<20jrL4~fmacLJC6K9=#<da!gjtJ<Guzaz^<kHQh)VF8lJ0{x_s"
    ">lRNV2n09y!5@uc_av7^hkJW!!WoopXQnDlTV%D8Qe*hNEciX=e=ufz)i>=9>kZ~txz`&u0E_WZ88hEF)f=U_)GKq6QG_@(IiVVH"
    "6!;VciRY3mlS=UmxzY=Xnsc?fg3pEQw;f9q=|Rvmc)>XCRFw4a3mr@<KZP6cXVqa}VLiT}iV@{c{LbS|=6DXO2^&pLOe#`uHPbg%"
    "VBm&JYrjpNmqEg+Db%6PBvw(AC={A@)H1x)MkiSKTY<s7R?~MbLO%`^O(x1_%~p|Ke+AX`Z#C=u?LNz`1^%1kl|B&(O~ac+i%+C`"
    "5dj@N*!LTu%j44Si9kr&-BxCI&=Vdj3%>88DhNIzygJ!R-"
    "M)=h2c5(>>@LgQ!4llmqoqlpr7Ujp8YrNn;B7wE$AY!Idj`J&OSXNE91M$@+sRyUENR=h^Vw}BZ<CL%ShhjRXIxwP*{cS9kWAy9%"
    "jl4;3~uU3Dhe{^DCBA@oqGVU`UE%n2{Z*M<-"
    "M$(TUqy7SPq6dv@{8{my?|^6F9T{>>TEtswtDUfua5;rT7BSe~lJL^r5+)O(RLG0TQ!&HOcPdUTnCWE896^B6h67!{qA_=oAyLWC"
    "_Td^Ge6#7NHehQ#pl9cLER^Mm8i=c{j`NVSjiF7j`BpR-"
    "&g`Key8KHr>v~NCrZ~&M8K{{$2o6A_gtzlU;_NOT&~Ua}B^b0Nc?>bFHuzuM6zC&h)DYXqdKa9uo@K_|v#89e@+HPbkC&lUwdgVL"
    ")giBpOCHr8$3wUF8dSTAx8HH9CB2)yv)U)*X)~(u1I3;HvqqzL2A5cHndoFYfQd9lg$6-^+R+B<yTad_hRQev0No{8p2-"
    "(d0pf5NVKOPyO;>*ccR+ukUK64n4slJev%e262`o5@6ltzx@n7djF|*pxyH}k%-VFK%FXEV-zmSh3kt8V{}UM3*qm9K3l~!sht26"
    "3=;MhY!%c@=qh6onW61d+bXI_ca+4EKO}kPv{g+rSaNhZ`(SlMl(yGEr9qV?>B_<Ik7NaPKE?!(fO;}w8oZUP`Tl8R!Du{x!ZBl8"
    "<uRG^BzVeuX(m)`)w}FT0S=)6FlZT%owjcU<lSevWi)eM49ID)ShZF&kuniKvl<Da>ET++ItZ-"
    "}>IvM$yy;%AZ+3^*A~pL{sE#nhi;+=FA|u0lrMu=UQ<SlUA`4exMMVt$6bmcnn>$nW#eqC;p6Eirj1MezocYrrplt9Il^BCC_D6l"
    "p-"
    "L1iO2q+u8m^SYUX*ur8d<Q=nZm5v}!;O9?cUP=RQ%S@oqN2=+skY+`E_$Fy_k)}n`7Rdm=|5qe4r?|0O}9B)10u>~lv8a1r@G~c@"
    "$s|MSpHAvVkwFp8n@&TV>f2?Ex6(K;J?psCC~VJpeWka`@F=7pX<HclGpLO$y~JgufHA7NFtsjcE+JZz?7{ygtd5a$r$Pe^U)!l5"
    "}*1B!&>k7Ga*cmZLX-QWuR)%=^D!=@SfaYeVYYGDyVs4$I{51RL|JDk0j8E_Mw|}C$m)p2Algp-"
    "2sv>s)K<>j6#o<GKrR2{w8zQQE)oHUcHNvJGyZ#LVy>`q!v9|!Zcc{D~rT*E}`b<JK6LqA__<GI1{ghURm!qkW`zcy9UZ!+Z)gG9"
    "zxjrnh4g-"
    "rr+*0(UhQRuu^6)Z~tjcjsc^4DCtLNY*nd^hD9#5F`GV)=T8HaGJQpxl7^<F7b0oHdC8915_HNMJjGH_k35_@F|B3GQhG_zI$Y{`"
    "2jowYHJfoB?&b+C&p+w2t32Gwek*dTh!$5pY|5-wZL+!pA?@Jnf5o;*-"
    "pF|y%i5|zOPND^a!*&Pots%f;~bY|6&F%gJ#L5jc|JyJOYj{ZH&|y=nbM=jqH<7NNaDcodc#Fs+29##7)_QB>lTMtr&zE^oQb?q#"
    "$i}&{=YTnGaixtHFGXUq60z3#)UMRTf7D82=1$cxJLZ_{xsCKq7cZKw`d+EBKMn-vcC96%j*nXPZH4sA!F*2c{CS@vvPQEl}I=rv"
    ")5z|*HgJ7pHIfHWpjz<7r4AbcvdgfI$54**l=5^ufssbz7=EDWhhPiR(n$Bx32HJ^fCs4jCoCBQ19Wn6zAAu@N8cJ+NfH}ZyxyTU"
    "Dk6w#B3ZGG8V5HEWRZ-"
    "^aBwbi=I@vtkZWqHcJD5ig6bg(mTaGxbc%3>gG=IO4G4;oE`)j8>ct|z8RDn$kj1cI<W!;6U)Bm`pB1vg^a;dDF2?sTi6{RI!-"
    "8A(9}$}70#B3&X^xj17o~p_r7CIUp)#kMlPl_a?*+*jO=<v(4!z@<Wl-"
    "0Gk5!C<XVQj<7ekp{^ho77ugJp3<IVE9}KVO|2&0L*O6*<9s<ATxemPPQIIjR>C|@!-v3L8yo+qr7tshBjV`m!!zRzKd0n?2-"
    "8eX8Y+f~ohz{WNwXar6b~@~V-"
    "*X)w8h$t#lh(|x^0pMx5`_?Adx{0?>rH6y;6>z5T=%sbKv6M%ej&|59h|;I?Lw$%b`V>ugX3Us$?=gp5egY&8xDnhs=}v1MS3&_3"
    "M$r)w?p;g$=Ee-mciHoVB2si`w<EnzK6KN3)R7f{*&NrC<^%;OJf*Nks+Xv!Qmd3&cyv}S?4n?FFS&TjJ1my;&%bU|5AlAdHn7+Y"
    "xJfTzZ-%eW8+d<WgQ_o9-Dm*jTN5T*$)XI$k@1?p|@Ua{yX7JJ$<k=@Izqr;L(v;wAv?PoHiqt8tem#D%s=@2KoQ3*r-"
    "k)<op`5(m+s=(WNAVc*t;p5M|^Cu_fwuuocXJ5gjnPjHEcg4g*Z^&n=BSfe3figCnftb88|NDjP?sXsuj1VfTieduG+hBzN<YJ;8"
    "7|6={#bp*R4depMAK8M?~W^g0sc_z_9*n-uubl2oxMDLt$)-"
    "j^fe4r9l{HD1R!LtW_vkYosRFfr{Aes~#b$|yBuaLn}E5BeQ^9sC;LWO%IEY--"
    "!qi&53F2p6d+F?r$@?rG<1iJo{d{5{ZCwlp|RWEW0Hnw8GuJvrQ4zsHb>!X+Z6^2C)Ft#x$|qPwqP^*M~kA=dhMW`b(alHtsfLQT"
    "|FNje`dHTv>^jAE-Bhnft71<U@57<%<y7iwB-`q-"
    "S52wF0n6|IZ$1*cee2b~xnIxW5tuw*E+twbZ6X?PLrSxkn8H}tHcQ=w5YlOfMiq~v5IFO<-"
    "Ub%(V(gJvug>?MLG8?7%_tOdTrD1i^~sxM0HD>xbiJD!Io{-"
    "t6i^Ym5iOV2{T%jK@FEMle6!zVr)!Apj?nq}8=Y@x*+hRn4jnZv0p*nil@79gQ})aX?zRTlrD&7Gvpjg~V@qvli83!`9*YGHB3l7"
    "er>B6<{@k*~I|6s<)kCu={!5~ZAnK9!242#T>i3&pcXvX84$wv~L#+_i`#1xLTVu~M<P;iYQt#k1Fu?E|h<?WayNcEzJhrqXVouU"
    "Lwmq!*Jw;j7SAov1F@CCx0ef{HYndMwpNYlnnPa!y#z2IO_0s(EWAvZ*xXW>y<ZSa7aduz3>yWl*${R(P@A<O15c5&cC(NAt9oLx"
    "yZi<te99E&c_3qdG3Z7Kmrkjs`0kx=OYxol_~yBB^&KD@9wi*S3>Na2ZSa0##*ZOS7%hHS*%hfX}Y8cVe=tZdw@SJWFx=egy0M7<"
    "W>Jt&PE>hltL!*FuKon)P^@dGG?Z2pP=@4VETYr<AHd$JrLGG6!2TsQ2VmY{R+}^(}7&z?aatmKvyJ*s@fVl5*S+3c3i-"
    "1+q0=TB)RZ1^+yVH-;l=-"
    "buzwhBycPJ7Y(ep=;owH{;CKDCvD!NEI$T5yIA*EC=f&3c@r{$*@(nmJ>)b@oQ!&0jlG`YEEQL#*eDOplL#Ij?EROJ8K%GN12PwV"
    "+p66!dRABv6TRl3|m)-Y(QfzU+L4&0lpbnFAYl%W4}t?m_itHIT&~`<s|0!vSs!dRH22_6fJ-"
    "J%_HcjXB}?})0h+S#O%J9WOwshn6B;0w$C_}IRy#fiZz0Fii4TwgOt2^&YyrNZr9hbUiX_#w&8lvpNfE#g|kSrpUQFdzt;w`@AZS"
    "Ti9a=Qb`=dNvlmm{*oW{>(T?#(U0&*t2YSLbyeZt^8YHCbom#IwS+Ej1i9xh-Dw^#^QA*6_E@xO(-"
    "f7z1U_H6*cwmek1Svb4MtluaGIt@YcAx~5cslvlj>lYTAdoV#VWO8D$D)@3gc`JaZtr6q0#X(>4Riik^hrs4uL8tsJyHP5?YI&>2"
    "1;f&j&&YXDV~K*y?jwTR6RW?t>o?FOKv^AacD@HeKGws9uD+Em#4YqdbF1w1Svb`Gq@bD<<UXtSN&P7a0K_`MsS?D1C^=$sktF~V"
    "6))CmY~+6X>RmzX?+dgUyjxu#dIsYfR+xFOgVlc%&2IyVmg2d(dR9pQ4}xO&ehjlGoV30+vKLP(`SJB6flD(IuCVUinB@dDd<08o"
    "epa?`%R|<#7u~2(^1Y468NZTa1Oj6$3tW_2xyzUlED}HNmDN?(zX4@)O?|zl1pGQ`$ARY;vQ<PZi`6_v2L7IzwLN!j{zWM;1nah-"
    "hw3KOQpQA(Q%yw8h$P%iy8+_pK);EgSiX)(0ld9h4r}JbbRY30wHDY#q_78OVQn-"
    "6ZinBfhnDhMqrS#bUu}%eWVfTTjr(EQ*pA7lBM@Bo7Qx)d<!D0V#v@5zsk$+W#e>kGyO$vv07<*d{EL|)^cj{RH#VtkRq9R$x{Gf"
    "RZXyVo+_D4UP|M;PGf2?kTSFB{P;$8D?A&$nMSmR-$>=L=)SqqWe=OIwdr-%*+y`xbi9FPgGq{^>=Ua89VVi}ZC@k$*<}w{Z-"
    "xV_H4sUkLNg7mR?-khkZpUXK0#qp#LF$Q*^Cw+2a_~erHai-dwXgP%`$t`XkePMw>P`mnlYS}#TdPXB+l}S;qL`!4g?3G3KdRWv1"
    "7;+eBIMgxRkcAlD+2lk5W-sZ04UL#NF<*oS1QmXVjhsFeU12O}G3>>Oo(qo$o76w%JFb@UNmI#nhrTA(6K-"
    "KX4h>(1n&<M6=LUtfVNj)cyVpzK_-N-VD~k4J2x15s0KXDp(W0N*`**qJTv|ChcNyWkSWE7-"
    "i4!4H%+ZUTA5CWmIKd<2VXgcr<=v0$x(gSsh?GQ2lzLX`vA1biu=S6V1byh?*3I75lnZx^{WU`mLbRRn@~KMO^hC#L*<qQk>)skF"
    "0}=V)XFF@XXeZK%~S`#ook)GZ|bzm9jVLJiSKe@g~DF*DDb-DFW?HUr4OR!s0~IEZed*sWMeTA}VokqLqubMphokSi?uPm}8mcI*"
    "<DdQ&MdRlN41B;vr*;s?h~QXjWfeVHV{W`Lgl7ff(e=h&B$|noF6zeHPbj65@TfwnkH?Z8k?i9{sL#OT#HssxY8%_KC|iSQ<W=R+"
    "<Ef*ez)dlj?$#f~;iK9*p%^1Rpg^(<jxr9)}QW&<wR$YP3)eA{hcE#a`aN0{?fgcj!iD-+-"
    ")|hf&ZyhME+M1<Mn<(g@7qn$$&EiO$xzNiiq<6iVyAlLh+E)ratQKX5ebegn_Q^JLVdSS(qaJ@Zk9D&EIu&+ur9k7x9Z3Y8RNmZB"
    "=zbWl++Hy47bmIhOXEo4xv&2y4_@hWBB;>L2f7@9>>51ABs4jORMX+NVo_)!*#wl#=SsSXx%4c=?&nSGSi%tO4C$g?$y^3_N>=0q"
    "N%lweulx7_itgk2H=lVZ<a<;n2mJ5uMdtHh^0G1>;p9yk^^DIW7S-"
    "B8n<2}JHE*?|<c=2mJ{RrM0D=v<c9B&$`fFt$>ydKfjvnXI)sPEep~cl}w`Ka9glioTNlVk209;se4&q?+NI7aIjkial!+EhAB1-"
    "luzgSn?Qe##)vw#@BPaNE5oUW;UBI1#(iH+EYIHEYy_3l5qa_VwB%G1oq#lI%vxtXh5w73u~*{_0d#vrn(GL6jyB)<8F&G|CmsB6"
    "ICtl5G5(5Y)sC~Hn(vG(=+E`W?(+u5X~?;aTIM&^&3-"
    "|mpSs$F}vmy?5cZBGVUvzFxxCltu4dhmjG>gV=oEVHk}^qp3L0FBw=s$r!#mVf&i;A3+5a3MYKC#QyeN41)NKB>!Dz^ZP6?@D&}*"
    "%hzW|aR&c^|FC!ZHc&uhLskZLP?#u(8qzVtC(Y5pzKP)r1V?GmMV6|&G%{s~xxbHNI!@)Xw+qL4b1^`xDn&xsqN_=DhQY?kf2QoP"
    "|rUn64V^)nbIm0InvSW%+L*NeQ?s$#|a>hO%t0`+4Mizd6nyFZeT-"
    "b26(R8h(pNIviagFmJMIFqRRDTZ6bbr%T@6@~OQQ!3(h(s({jk}o2xbI}=VBD;CoV)oN2w1IYIy3zghL-"
    "r(Q?;m*yY_B(diCS6nlZ&`>Aj4oiskHwh%x_rmn)pOJ|$@&U^Qql+nDpu`i3cQs$<SSqds+7>Mk56gWkZMtt&m{$*UC~QRQheXry"
    "G>>MI;b--}r>@ib#_w(whs#KQG}v9e^Ti!h-f9#1#?_$GKHWaKSBGM{CWNc@rUI*-"
    "sjFcG*|IkKsd#2589qIX;+XvN3U@}NPAm7Ai?=jXkyiJj%X>0o+l-"
    "eb=cRP9MWZ_Ajec1#>U$HAaDEH%;^xL7$VJCT4$KG9&5fWUq9vD<Q&K5mDo>5fi-"
    "?+Pt;YPoaPAok2^ovB5*tJai?%n2JLGDTD}JGQ#x(j_MGtZ*bUn3kH~(q-"
    "Tgoqs5#>TNZ&WvDtq!4S3b6kBa<DT=JH3<;o6ja4hU=<DUzeWLOs>k8NXHpZUb5zJV*DmmfGR0E2UEA+*pwX)r?_W0swR3UBrRnq"
    "o4$--{BGodLKC8y7qIV|pi1}au^s&)g~@LKh;`8QngY>5#oB{gf^w256tGwYJY-XK#**S{4B?e1t!FA#{XbK!624P6OXV)iX0HO("
    "1bZ@8!{8$4ULm(lWJ=hqPH^xD6uGm%eQI{sRt;l(47H~i_8j=xUU<icuOL-"
    "8YDM*k)cMgg|JMl~xSQ&M+Z{HTL4ZS($?t#9h=do5oo6|XN8t`~UgZ`{t;@(YWL;qQgLO|_A4LFMf`SZI$SH}^6=zy4;Eb^V50tL"
    "uT#HN2tWjF)-51;}sVV$Wc2JBr(WrRi-tUcD2ckXcpAW@G;z*2-"
    "jKuliFmwH^rF8{O2kCa2@5YIeU1XB;V+yj4#M=r1|$wM+$!%p+jxcRQ1g^o|-"
    "L<*K7C*L}k3$CI&Z&Fq%{UbqTTqc&R1waY#HT(YZ&vHfk3(b&<9YT(J2Ixl>lN{9PS^5_LJ=3Y$ce)ElLTb|`kVpC3pAKvsqTEua"
    "K+&D;diO8qiHF?nEGfXQxQMbu^-ujew$S4T9Mo)4091IxSfneTqbzQ&F2c+xF^*yf#Lf7_1vws3bhur-N)v-ohlvcCr_-"
    "5AtpyH$~mdsAd{|?JFE<puNak_ua{M=plF+vRnGA1?&og9OM`wBoNj7Yi5dOg+=Gg4gM$%#<N7`u`xc=99cO5aPd`z<wil3NS>L-"
    "&Cb4Gc1dj-48}WL|AILB;EHyu!wBv~x>d$L}T^yUl<7?Rb+Xe)RbAjL)+<P>$OfOV<~I%`bUp%!j`h9MrOL@C-"
    "%(?7@chMPZ`R%LR_%jJQyNk!2*mke-pjz;0f*dIW!Y1}^gpIeh9;@nqV$`;1Tc^q_8VquHMYU+G}eWpFVAgYV#Ltlj;#FIwGa!A="
    "AcSr%*!gEtU_rN<2?W0_8Aw&uX6J#S#C@+5pPu{8)@8TP!b8ITm>M>JP80W#U=&z6)6#OwHPgRg)2dhws?Y=L{;riVl>L=TlNgNx"
    "RblSYr0FZAUByfVW>qne}#Ol2Xpr3v#+imb#P7S<bjEVpLAct&87p~+H9SLy@{idAyXGBLLFV<t>jA{lu#hj~p~BjVGg{c-"
    "hN&Ov>EEq$87awX~tmC~55$A^t~7_y~xh?OpX70X&1LSy8;<{{*D?yKxNmZvZgF<maJ*5;0z3jeUo8z)&IXiMMbN_DcRR?@jBsI{"
    "bxldbd|%SnFDYSe31gEFamHW0;S2sd}w(%Q+?>4aiEowJrGTVgzWa<EizeJ@AmMxU$a{#yZ|_gQWkjaY>NIbCKghjR?gpQyq&^Ux"
    "(~OXFv-"
    "*PN+mepp;L8L}<$Y&OUJiFZP4^X1dDu+L23lyG}{tzuu_`E01FAm1vUm0tzHXBpduVkIPk&wBV|2(?!ba~f3=yHE?3?}g%L{T7<C"
    "7(HY%+*uP4${=$yvsII{=1MW8+|Z%STZ{75V+l1H*IZ%{o9?m}Pqw=k9vX)#MnEx%fK>02;v%n1MMfc_l$;1TqA2PqN)pW(M~3!M"
    "F$^O<kEqTXxm%76ANgp4rY)Fzk3_Zn1%OTt1=jt|etSxyBI6Jd<iOM{x&;?^06XQetRd9Z$kkb=*9;Io@g6$XKr>+?$U=&)(ouQb"
    "BUo;=lD_x*4d(+(^-u_!y=?AMKZdWdQ42TxE_={#xNc34d^$k`n-"
    "q)090GAm2?t_Jm2$n=4E!0V&7}z_2%5WUmY_R=Wir6`A94oKAuONc@@(uWYXw+0np4?iv!Eo%kwN8Xq>UWD?h9W-"
    "nU1s(90^&E^Wl<+L@4ykZd5-"
    "S4D7tr%tyS{@3>!pMn0UNeNA#lM<Vwon_q<sEykX|h?x66<PdpEzw20T$^eNV33;<Ke*yQI(5;9@x9k00?k;mIF{MF3(9R}xBk&_"
    "c#4|p}hC4JJG0DcIr8OAnS=v^W?Y8W}|4x=5AH{Be4*jl&rh*|6GC>xrmTRTA<(7S<sYT&ARo;P;yN`y;ali@ES+hLSJ;~+5HOZZ"
    "<Z&tI{>YU|df~DtqUXC!frc~Ms0$?fKoZo1Ht%}_AaW+KbN$CgVY}MhK<Iq?NHHinT9gWL0<!dmGV5FSbu~=oh12K!YLa|mkF#4@"
    "vC+y-"
    "WwN<eaB(7|YxCtr;u<$rvB?gu%bxnmfD99sTkz=a{cS}r&g};2J3DMkRyuk!bw@?CJg8bR5!=0r5b7mE}b5(L!N?D<MEU7&LE4>$"
    "*z)=}=-^=~?-"
    "W2L`6S2gsUrMsR`7Op>*OhIau@pCj@xwaY5bG2RaOsngH<Pgi8iub%3IMh49oua5a`&9qg&F{nF?_*1BWhoIri0lBuT;&;FaUhl{"
    "kW(^Fk~!ls=s{(S0xG)-oiz0$-7J(4Gu+Xvl3fiy?&Q%G#NjkdTqyJTr*%K!;7IJm&<R8J@G|2LkN<zY9jJU@Ce-Kciq;@B|;%%?"
    "6SFT@c?q+J0iTHrY1Um8zM2z>k?HIWGq}sJ0o{r)AtH3<!epH<L7i3$k?}<ZH&%Xk>M%M(Ro47MKRUSj&@hR)up=lZM?!`w=UNTJ"
    "qjvDHdV{=Z_XgLGEBEBGg;R~rr5F5tpP#C#*5}LeHXB^AH#n?BpbW`gmpU3?QH@?#_;)UNBuK-"
    "q}$aU?^pYdOJ>KyA!GAG2LIz&)AkQ8$m&!!M->Da8=H=jLaFuHa3WP0<@{sB(%<v~JGUJl-"
    "N!*8WA9R`b>s*h);;xbxa2py+&%Wlo5I==fk4K*<+O8~&t#eO9w3Z~x32o_M~&We#y4xQkTH109Euv`1J2b`cyMy*JZ`!VG^q&4S"
    "ht$3#d})@bzbVKhU(2HY^rVE5jbRQp5?6gU+M}C2gTdo)Q*Y&MQeS|n>s4@hm+OC%rO4QzLS;YYR`rk_Rjp=9oG(Z3SeY7$T)m+a"
    "Id|FXZaMqxs-wW^Z4z5qtAH!nr9MbHJ#IpdD?8*n%Y?^m#kHHf0in8Kf~Wg^Kr}GlSU6sWw!<@mFKN&X=ux#oSgADrJ5%98ziON5"
    "rkyusaTV}-"
    "i4yDJ&n$Fus5*=BAM$5Qo2iF)pEu5iLB&1j2CU+!BQ@&M@yH)+CL_8(#)z=7ObtDpJ9`am1*DdI*+hS5LZ!>VQJA?chX2kL0eQ2W"
    ";aAp1(x<CJ!mrQ<*jL1h_2>SWoLU5oSv~rR}2x7p{MW<=aC3g#eevUv+@-u=B#9GT!RP#y^uqFqR-gvKJR2lBmAO=ONKN{;g0t*;"
    "CdKt<6w&jZ@+2hmauGJiEzpAR<Sl3`aQ&FW7)nHALa_984DeONrtAX3wK_|$#w2L59cFv=0_lt;m(o<)((9CHe7Ucl61jymirBs{"
    "Y(Rw3S-"
    "p;Ynz!68`^;o`HmXs_=c+_<$yXYyO{<r8P@F0lRbd{UWwnTsm|OQlHE*)lnhyUYmZo#LPac(8JdblYXSXg;U9!jE>W!(El(RYRcG"
    "TGI!dyPyIOK`DoE4CfbnQRvfZI-DR44T3lNIcb544NEa{J@@0Y{!2oi3rv!ynkzL6A5sF`rbsU=NurELB3cVL&p;I*16`Zl7}b|P"
    "ph1lkfA4{K*+2013G0YyLeSQm-<v>q-Q(iZKX55}oSpgSHAlbNt32cBL}1&hQxRpAwT3gBV>_`qE5e}i>$h)Vv=(8-"
    "W#C!E3yMR(z=Fp8=g#aGI;;2}iv^l8wNA<R<Mi7gQ}6nxI1TAL-=@a*!2Sn40h>mb(b1zNHzgJL-"
    "Z&Vz)gB?)nEi%a;+Eh)+TIqdvbp2m7VS7*6Z4@-%J0XP{R?MaEJ7pk6_fH-"
    "x@<Z?>QaB)xmRo<2d;MZNm9x1lYHoeXibi>EtiJ88fWcpO^GWPPWZ2pY$;pun?lXDX96h+47hL?P*vEPT90hCMrG-"
    "iGuCOF!dZCWr-nLdQg@*E;IXK~l7%~3g4>(!$mWoVP2+c9jx;)#vbA>Q?cqKaEhNDTy1CN8C#nEqViJ0Db(@SJwc6Mfk9I;`vd#x"
    "_7A#X#A-"
    "a|(r?Q%!lapS$O~m*0$hJ}KjxB;CR>)Nrmj5rd8+CWaB(YXLB~@_Bj~q%2)EN{#X3!KpkR9K!KnOYL3ra!U=@!qf>kNLjmP)=n0!"
    "cW3bQehg763-&5&vfd`2-_&s}oScY;lG#nN;9+~nA@~HR+HRI5=jv`645VzlXq43p+kbb$JBcT>_tn88-"
    "1F=7n!aNpy#y$v44%&*5-EJI-GSh1ydbR3g$!4xry8k2M9TUF^GZHO+}FLbZyG^g%=Rs&_e(B*`g-"
    "v<_~&liFIj2!8;)1*L@1;TUNYa%BJQJ;wOS3fNY^u9bO=aUxST4c7wk!wYD}-*>`x)47k)A+o0<-"
    "ad(yJcG?fmw{0@k2IJK<?0VxZc4v7CTl3vp-"
    "4Rxpsa=tj|VUV(PErTfX37i(*X!NwW@XreKXCiy7qMC*IZ8edW<3izM_}G)fr!UneZBWlz>UU;Bhzg4~Az4(z#g?|Mk;hsNRSsss"
    "hl+|eANllr48zfH;k@@j9t98Lp>bZpo#WaDGhk>ePTFi0GH)*ju%~{KKFk&*B)P1Vy5V$qw$5+y#-"
    "!P8JLAyOX0DilIS3<$tynMB{xMkh?$3xN4KmWSlw#Y_*~sc16H6Mo^4ug@mH~G+x~%iCnKNfy9>5>!**gm#nXrs{pRaWMlV8hzzJ"
    "t-rNut!!MUuV^%JypYMJaSxhNjfy_uH8CAvItr9kp!D!w~M=zQ`h|PUadKp+OZUDXJ>gq<lWZPks?EkFfQxxwU=+jeL)alN4Q+%I"
    "JlHLvr=}O0$osc3zJb*|t_LUtcI*U#uo=YxzR>dqHH)vheXB#{4NJe7uZFcTf+Q(lcmlD8}gEvRLdKCahuE;50}{5w>X49T^w%Lg"
    "bFjHV8H`RhE<tM95-Cx+c47M5VQkMR!66l@wzIn~KM;wbfxad|X1uuL;sr{Qui~yWTdEEL-"
    "%i5c6Ia5+!~;w(Zl`7qT^ir5?jigD9J0@=|09Eb<oW;Q&RoaU{E|=FFLv+AX_A(~^#$YI)qSU9N8F-"
    "_%dJ84(#IGcqzGQreN+JNp2QYFU=dm3w7IM(n-"
    "z+F;gmGHrXTy17i=E}KmFz&}Ui?~;j}6q8^hGiK?c*q}0#$S1(Ec}&MlibK!@Nt&bu(QbsG_62HMMouLiw&#ql<PK<#hLhA3%Nvf"
    "eCbiLmH|TsPFD9wSzTX|-"
    "^XHtQrNkT51dW>s>d}!%TzL;PStD;h3JMxa(YOLMZA1BC8Vvmrxph+G1~ge>Z>U3*Ga=h~2RJN4I$gd|E_;c$TIJ4o*hMyOD2A34"
    "Zyxd_whnD{NfNTH+)$%3-1K^Q*2%>S_-"
    "@P?P`)IG_XG;s=$ljkO~#mwqBb4c#MJD~F0!H7M9*j~Ff+|7q34!YqK!EphYwh{kH$SyM@osXTTH<)Sn%D%d1_%6n&q;k0-"
    "B65yPt&@W^(TN+W1-"
    "JF+=w3#juiM4M<huXsm~VHR29@Py;gNGPBnc(K~A>25hi4zG}c6TvEir4a4}z&q4fupY*ixAUw69nSv36Ns6i(*vrh(*cHqBjJN}"
    "A3sgRm=5mZkK0?qAxz&KLieBoCHRA2H8Q?2|Uear=ftpZpqy~h`qt2;*C<~viyU1GCbg-moYe62Zb4)>4<cYcrCj1S@*%X%ot9-"
    "Ct=tBuNSucE3t+O64O#7cpVZf4zcq1+1S-)(*qFYBh=O~KIu&eR#^H>WULqVB_s*|()rD_2G6xBo$nZ2*Zeb3HnSxCr)d&Mc-"
    "2WkzTBvRmaSg(L6#{iHCuFXVLdi<R6Om8$L^0UBsk{B#IV3qZ5h|lj-iYz>2f?l%+UA)esBG*wO);vq;d5Dbx1DW{RY(%FMy*~>l"
    "MA@4ju8*H%f###gM^8I!hdYb-40RS9GBI!1ty`lfb<T(Cy|`2@bs-m#b=I%^#`~P6EJ$QR-"
    "*md&kHjq|uUUTQQoG#;ftW7sc_K&_8ZseoIgLv=l=}g3E=<-3d7WvCRo1py0aFtUP^bjkI-"
    "6pC40(#6misXHA5f7BgiNSyBP!mj+vt<Xdc0R2X%~9I;8(qH8TnOzF!pR}Yym_j@Kw9F{Do?5kcGtq0rkMp9kt267I!)?iQemaAA"
    "n=QMTP?Vyoy;l{=i@>y(#9EvA}b@O(4j`c-5Z4I-K23Vvbq7a2$qr{dv&!ovJc`AQNND=`x|5R0*Ql(K<!-"
    "x#M{dfEzM0JUEO$+*ilG@5MEL7HHq<nZmmrHVnAD#s8$>EI?!e-"
    ">^Rb$Er8xlXz!MWWPnSQI6X%uy)}4u1|$SCg!#?s7z`c8*I|JPvM5|5hMD?lZk7di5K%-"
    "LocL!FxlhuXhGiAW^hIG0Oe`iaJRNr)d3!$GgSfOhw{n|6E(b~@o3zR{OTgFYp3i@v3A!$7f$vTfilBThGk;oZ*o?{<*TM(Qdwf5"
    "ye|%w=yz8<B5&g77O5-"
    "`c5rZIz)FU$Wiad3J<+Ops88z<R)t$wq^>$nGCWnmted2I^rN_n7vgm!i<AE@x6JyOMb_(exEB{2jlyjuLu2e@zYH?m!gO}<zkQ{"
    ")7klE5FQ!`s?F8JFMeiZs_n_!MwO}Vh<tiAJ`I8w~h%7_a1IjDIIm?Y?S*4r~hjPfcO>qdz+p#27Bq$3FH4`_Qw)Z!KLo=VvyVi+"
    "hyck$a#Y~30CfFpM{8e%+vvtZ)PKR>W$fYbna$F=XL1`^PX3d-xScp<+gr!<g-"
    "WvH8izxUjgR~LKXyXPNYeXE<7Gpy>Z@Bg$pdgNL!vd7y#!Z@auy9Ads|V1E!`{w@%T<bA-"
    "y4+uM!r^0M6A*7359azP;T)&<CKugaBurZtc}apehw)avYKEiu|09ieh><x|J0LV&npkvzhFtE7{HR@>lVqABUJ`-"
    "w|ux8hR%hir-"
    "{Ex=V$hzCG&u2Iy!Dld63FdeblK$748GW4WNI^4Fi}&xSMI=&iZACOS*Nub9!!M20rrKh=VQAbb~YzwhpPc%DelmVc)#5%_<w7Z6"
    "FK)=tS4n{4|}b^IljC$tCwFv0wVlsMGeV<W2=bC)TU({j7YucWf#-gs|Ghb+gXFpc7`B%f=s-RQt-jT8H=ZsDMQ}1qYotJM=lp%+"
    "9rWH~6SM7#HL~Wxzov&JGS6QTtG0l5BX2^=zau`(!!+we>8`cOmX6te^B{g11J7htFPT$o#GoS#anCz2$Njr+Mm!_wr-"
    "r|5mox<YSGk%_w8mX(;qoQOf;=o~@J<0=Z1An6^g)7&;+ecDkKkCy(HjE#FpYv5%$`(W>3eNG2YlsZF{h;CbbVZ=V+f1Tyi3($_d"
    "0w^5{GDWA)M1U30reoPZ$X}=<%1T_vDBSe1WG<Q%?!kHVK98hpqiX#S;nkM@^jptb-ckl*M%H&EcVQDp^oPmN8&E!eEVwn&{IYSG"
    "jmdR0<z}5>vv48=ko5@wWLnU^iTjK*9X_G_Bv9NhIgoOnyM@7wvSp6re&G?;c6VT)xy=RO&nfaI0q@qcLy_FVr`#+I=Ww(xdPGyY"
    "~0YYChMO&b>rVNBsb6EJxR<%nZMtx~8YK^)J<S<>oeyt9;2$5hD;;1I;s7>%Z(GtGp*-"
    "Fs>fKYf>?ctRR@VTMaSBQ)sc(;)^U?3FXn$u{6S6ms*|M+vFCHHBEeewSkFcKjKflz>L(`{y{SFyqTt@zIOdUwNPF@f3d!kmSKP{"
    "14ZfX`1&z1EivIpTHfdDSqZpc7^&&xSMcV<+OrQX2ms1^wPO77L?PymY98swvB6>%?o9J48PYuw<2DAdBY+%AU5>T87UGek*L#63"
    "?4R3}SiSQt?vevg&10`L(h(l;8iGd3HH}v8`kAWpuj)%2cm-"
    "N#@Iu1Ic{;8rpRJ3^Ab><tn7Pku>Z2i+C9wC7q3aRk=UL(sXVHOo+G|U=K6@nbTjc_jJ}odf^F{B8&klA-"
    "?L651KVuf#?|{n$>6l`(|dQM^U+DlFXw*)uwM#Qoy3XDoyPHuxOp9TDq{RHTjxdZON^8t((=$MC^&45{_K<AF{6h3U($WJ^r<&l="
    "!>-E4C}VH3|7D*^X>P#q2rA#g(&5#g474Ee+&P#2S(?*}ltPIa{5;bMoGRfm)a?g=>Cf7ly?Ai+!2%j^FRB9p3D$?bi{-"
    "jDlLM9meR+p)}5Mh>&@nSDcQ7TF7fox{On?|N9UXbNIRWKs$ki`&QO#j|z}zn6Xfcdp*apa;UG~d`;Q%ZIohyKrPA*$0$d}$jO!a"
    "QMZ_Ndkz4#7&oo+<)er7L|o#YKG!VpZTMpVKrO~LcZvTn^o;-7x1Q@iQoF>Fhtp~(F?}j!X8$y?Ec~il%&IW^d}<-"
    "K8Au$fbke8ffbyZ=&0FHLsl~QxUvWQ*6y96aaB^rk0&WWRLS-"
    "xPzx<}6p%(cS*T_%A_f6eoo1*&M-|=sBmkNej)K^`j{#rG5pC&@S=~)wM#y~C7HM=DAlREImZ|nPFw-NXkgT_9eT6}F9=fB7^cRa"
    "Fz-&6JMGpR+h;h4MN-#L^opG4MT%ky5tjDcFDZD%~0<emI-{X0W}^B?;?&z(|+kEa&lmi-"
    "+k=QsUHmuqaYA=~MAf4jv#omh0MOZLR_#T?3sK&e@<31IiVXWdevP>Z>(diX-1JQe<BXLIv8&-"
    "A(492V<&?t9HbLoM=E*Yw7PA#KFd8;@B*spUBY)PlT{BZu@}mqhX)|Kp&r9MXHU&JlZvS6ypUqNh|`Yhr)*pqTAl90IjK+m3w`ak"
    "xEK#-P7fGeE;{{=0PRipBl^X4~JeLAUZx>{*8m{2CiOL8Hb;;+Qy5BDK8~-"
    "`E2sQc@WR%>wf=8*vrFpz^Rg2!`p5^!Z=^b2cxjV6uH>%@7wv4N%`o$W<c;R|0VKK>eqoq3mJMtNfZRgd;OF^6WHou=7T}E>mn0?"
    "gWuZ!Xo%FgQL#R>Jn7f7mvt9xaht#Y-;a}S&u8Mqn@vby7~fpzD#(Dy{1;-"
    "I&IyaJ1p&bcpk(XQ&)v@OOM8PhR5e0bfxbG<F;m9EN<xuq|`VA>X4mQs-(-"
    "V(8X{SU9wEM9&Sa@w8>_{rpDwdFq3l5xD+|b&a1OzslPL1gZ2m+nP%kF_yjV3o~CTxiU$7|iQ@rlS2o6cAA`g9*^HO`p0p19bc*Q"
    "kA3iq;6GB2C1nn-"
    "e=TgT^?TG|xyNxz@S!N+}Yl_DxBo#9G=duaA%JMNRa#xrvfNWr~WlBLyp34^4sDxKi)ICkb>uqkLUW=ryIt4GW#kIBsMxJH}{=qy"
    "}91V8{xYW2y<ZPt0V^>XF*^OQwQTbcGO+Je+;_R%2dzZ)E$^%xp6CNy{b`n&4Y54ZshyEl&_)&b&Yjp&8qXKj^4=iJCw7|u;POdh"
    ";SgL(dT}oRg!@rb&FX^hf(ls4%B(MLK2<@ppm+5i$X{U>WuvEa*-)*%ON6zMQ7-"
    "rovPreKp4!%;rl7GNfp@xz0i5{|^vpelA5&4ZoK6|Uftp&qVd{Eiw;7M%TFM-;-t_C&|RP5OJ7W{?5PjI6v=tSe>YQ{^AJfJL$N-"
    ")NTUCHdmzK{%u|NSY^e^!yRy2iXy2M4LH*FlZu29!6`?4&FcAM(OJB#!p?yKG_BOz9Y@(bfdZky=#kSkj`D<?l5h3;O?5yMn){PR"
    "<jp<GgYo%X(?XN{qPr61@L)xcI!19WzPKYvBro&h9#fm>Pr2@C?2)nMiW_?qhL}G^3?PT@|?AFcFp+lf;hX*R0(`R&1C-"
    "rAFEc@Nm%|)!%+fOcvc_+YH%Y(HL54)UCoVl|TG)^;g4*U96THoiQ45Qz~X^9Kwx2-xn1!aXr`H&1p^KOx2G`-"
    "@o>=ZiO$WYn|a3PszyXGN`d!hu1rO6wTX)N<70>iI1bDMjena=9S2H@R?BwTiKn?7{{U-"
    "!$^&%Ca^$TdU+q{0_}SYizL*5lp1F(u%y^@I{a}wDK@Tbb=o^<l43EK#Mo+h$gNBr&eYpIsV<XSPsVpC4CC%I^%AJjy9~TX$@EIV"
    "wLu|ks?Y#Rx#YJNK|x@NJV+Wqd6?74@Ga^O`D{4MKM^9q0L@uG6*4snf$}j9-e4$rqi-VCsKF*b`Hvm&nzrb8{X!ME59GR($cNtJ"
    "JrG}n{Dnt~HDsAl>!nbmx(=M@o+`dh61>XEWT^7oF|^dEYk(a*tX`~Gju(_oTN_RIgW{q3;yBerNgNcnWTAqMM#4x5-"
    "DxrPI;e5n0`B>oxj?7dS}0d)MBWy5_{4Z^F2TzScZDM5rD4cdfIwI{Lr0COW#C89)q)?_lSkaEzgY#FwR$LiQU2uH#Cap&!z&i|z"
    "?sldqiO}*=fH8=2sWz&>SIZot6)Z4@>C_)p>vaw8XNSsF?n&B0aGIoPBt^=9iA~`Nt&6<NlRFiGFix}@d+)E$(TcitMo*6U|F-"
    "_<q<Ru=93TKuhG;EQz28MunGPeoZZ;FhWRmfxSz6695Qt8Llbgpe8MGVV)H-"
    "6^&IL)cr}T1i{WWHONC90$<`9^`!H)<FRm9N*}1+g=zPqHsstSC(=ay0I@)17onV0MaHxZ(Mj}``#*hx3C<o7!3P$ALw8q+lxf;g"
    "8QlzdIK#gLcG|qIA$@IN55q2*JiIFOstcOK3oCTU1kE>o>)%Hv$$H-XiIqRadR@<X`h-}p$Z4IX<vI9c~C%S-"
    "ph!Qwv$kZsT!*~7ei(ff0R18_WkFWE02A3LlK+2kZB?f$LC}aqage^R7PX@5mXlp|Cq^Vgd;wOJlZ<L9y`W`&r>4dMq9+oz>blB9"
    "GyhX<5C~7&$yb|LD+zlh@!rs%wU!~I)p>ccLr9_(qXUj{8+>D&kCapWBqub8KeS9p8{^O_j2V6-"
    "p9?WBkM}Noo7a5<HVQn76kcjkhN~F_&GxIWT9ol94*x~toZ+eWQ9lOHldH!$~CiF3P9)A1UG+smK*o&CT`Z^`qfGwPsF}Jaj!>)4"
    "MN@xZu_XEEZh0}4MkFagZ?Kyu%XvK;Le#gDfJ`#QOY@=)EU2We<!;SeZm%*6#{+NkTsebQQJ1+F0uUkSt6;}Pu5bF)@pX52l1_0="
    "T+HhZQxfSf^#7eu>=@zxnDgfw%+O+S(<`oe09S-"
    "M8)?(6dh#rjxo?mUw=+M*9$}x=hwKAu_5<8}E#m{`!_r<b3ZQQS{KM&fA8E%~gv?xZPW*00&wH2!24*bUs<z~V#jbf{dL^qQP7yB"
    "ax)k`RE;!E*I)94UkcNpwc?xLCJYDSEnFAu$hDQ~FRy9r&_AgQ+&RBPaT08S0<fP$|lW0a;qr|JEF$k=KExjLPQSql?G{oo<%1xP"
    "9fI!N^7ECZ6me`YsIR7O=xy(hEBt30C0iioci+~~QhdeO(9rk<^y1W5Jqr(|ATP#u0H?hC`l7@P>~ovx`$_AktIU}zq+RMhC1tHJ"
    "Ar&J2EoLo_|>CaCCHs>82{saaTbOCgbTEl?|7q$ZG#`pG6d*#L{?m1O{no~I_%c0MvK6;qX?e?eC)%XE7e%SJv0F?z~cu(slvr9D"
    ")(;)H5|K;3e;Dt)7yNK3Ou4%I|AeLEZ&ieQgr9esf8MRrhl8SKU^0Z;nk-"
    "~3xb!pGHR<Mseqf9eb>dY0g`8CJy_x~_gw(zezrcy*Hh<fV*<u|T<rq#r6BGkVf$kfcci`yZ2{4K~8W(+nnhdg@@V>{ntrIT643;"
    "+O01VkG^hk7oTdLq$&%P!=OHkPt3|pY<0N_q*8WVj;=^6+KISJcHg!cgV!1VuK;s0Hy;*&sYo0VbtE!v@m;7Vr@6oTwP>u4G9R2N"
    "u*>8>LPoao<pc+ht}|$nU`_v(B7zv%oO9fE;8m{vz&zqeav0E$f64k!F@=a_8$g?w35XkppWc|{ce`K@eiT}H-"
    "5}p3xwI8=Q~&beDpD0wa4^wZhN+881%9kw7H6&^?e7|3<Q0MZGC2Y29K67oL-lJxerAjzq)-v-"
    "52|~cLuW*A>q3p^nI_VIs)`jwRMZBWPBv8Ig+pBWc6U33nhKtsi{EFhuEQa><`AyB;JtZc`f}zdWi*F7+bDw`N!)0FiFmW?+m-"
    "WV%qa#D6j{2c`1kI!H0!<9#Gh#E?W*<QP<IQNb4JIBelba%I<hwP7MgqN7Ztq;;&uhgLL&$ye<y)iM?X5UAP$-"
    "c@1tV5cDBl$(JdR_>4|78>~>5nHduFF<#Bl#4O(;AB6H!4w$)8XP^V&QyTZQ;`2E8YCNiJvS;nlLfV;~AQwdvR9|y-"
    "xt4>q`Mv61n%-t{4d-Tf<S2m}RxI_GKbhz9t+T$E%ZH-mVnB+0wQE2RvPs>5;!;Hs*{MRc|IarGT$-"
    "kwY~mYmqJYv{_?loI8}gf<v|$R{cw9AR0O`@xf_gvk@$+KZ0RMF@L(G*>Z8Yzv1XAn^AXK}541MavW_1CGCN^gQM^7HygqhDOT_^"
    "J)uQaEjVYl)Ki{Q|V7Cl$M5_wDA99MrUsRBDt)4%R|aAL4Rkwy&<k=UW6BSufzDwM@gjtnXilOdw1R-"
    "kr&$5H~87Z<gf15~^J{D^Ql7f$Sz<-"
    "h{bm#8<uZiR<tzPaxV#|&8`Y}PfPC$s^akFr*qapk2*m8F6F0cKYpUcgn7rZF;(B9!Rafz;7Qzj&g&VB$jq9|ZmI{M8;UQciC|h@"
    "B!RBY<2gFjP(C<m6&d4gd4)IVls6<n(cTHba(aNZ1I2OKdpkSFyaQs0m$La7hh!Av~6(22I6$3oNOjQrcPXFPW0Z^8(94jh;EUDY"
    "n*59xTrowU(b-*Z}Ce_hYfPm|XuiI^Cz-"
    "@PA?3b{!geR^ap)`%1X|7em?jPEfgzrfh5mi=L_`)B_Y5dY`L2HutBPA0Rrn=5^YY2jV@&dYxc#-"
    "Q_mK=(z(JGwaZbUi@5HYJJixj&1+K!9|pIlr7vL=SjDm8I9PGJ?oANFagLylnDAtTF{;T70U(Qnu4760B&4FZ2&XG1vUeeh#Onai"
    "-h++(L6W|ytZG&KbKfcSF961qe9nPgF2E+jW@QL=h2=C0K`IU8%e8_?9uz?AfdV6@Z6$h06;9nHa~&?RK9et4TBqg&3eAOHcbGCg"
    "}9z?boC3p2+R}gYn`B1<lO2PDWzT9k+#_o*i#HXLu}0=;0|unJ|Fy}xQ{0CkUqcR-`Qqu-;=^)A5JW`E&KFvdB&fJH2je`H-"
    "9vEJ#alZ_x4-s_-"
    "&lcL_{n2aI;J0b}K5C7<_^8cmgMo5~HjM+p`(fwUYbCN8J&wp3OO=#3+LsU;0U0g_O@+=&aojcl>L^F&<&D2`(+_pqiWZr}bl9lr"
    "Ze#sB*_45@QKS@$*7BExZ)Jb*Kh~{wYy56?0m)*iHw_ODzjGF(zT8K+y!ReQ`w-"
    "?X<(^>x_|%57u!KW6Q%hkeLu5nFt8*3GHHViLnPa8sr8CG%S-"
    "RN^+8eYKEH5_&U%_vV1VvJv0L;bezQ4s>6?$JLM4(n*hU9xhE(G%EcfPV-Tnd>bL4=Vv3RZlX$lh6@pRd^vrmPao2>+dhu7sM}40"
    "uvtB)1S+6;y#3*Zd89bL>Z^Afrj$~m9s)5R>7>j~6@>N@?1}Z0s0<IJ@RV@Qh3srat(Nygdk(J=r5+8Op7H*DMN&{xN#Hd>V-"
    "i>5wu5j%}sEV_o`l_@6aFdiVxNEG9aSm<_F);#xvL#vCQ51X;7MVg#S)C3;pNQtNV{T8SSrm)T`?YX|>~xV0aMv*tqtJsl?tyxFO"
    "x<&K4IRXM*}SLxe~|PMJ6{GdlHt5`&xCc{v53ftGW`2dGm;FSQZ@+sm4{dk4_Uy8Q3>ZqlegVyVg0S8>2z6jYXsxlwqtp6N({be`"
    "|tXSe|#dqt0@6a`LzuAw~l7Mp02d8s#z`qUU3avbi!<F&pMl1fhOYaN;Hc;;#JADuRYJxO-"
    "F(`yeoF)dZOm^An+W?^Yp?IqR_3{N042s=5uK~7e57V^1|ywzY_>kKwu7Q&AzT4MB5BwyUM*VqIW&_tE7Oy9M-"
    "!1ZcRMso4z0POA-"
    "&d>jYPRCwvw!0mCq=yPfW!knyc?6qw`Mv`!o3kJtOiDLl&LhfCQDei+vFQS1CxeJC8J)_gas*>Pcxddq#x8b^1a*0?LNfjrh5*zR"
    "`5{T=R<vVb{j>G%pD0BVZOD)iij$TTKO93{+nAF#my+bk~~G2-"
    "V3YIeiC7jbgJC&!FL1ZzTynV(fJvo#K7Zq_hcV~e32tuX_PnXDSD`M<nZHBG_>t32e+)~GE0=a6FN2YR;T3-"
    "QhH!hcSzcf1bM?;>0BVt6qV)c}+C-"
    "+0T7MfXp4Zw@6qc0Z8fs%BF;xk9OpQ*$m@h6&+M&H%}WXDhxaRb9UlcV%+!U3l8Xc6QPMV<u}EY;W$C0>-"
    "j2r<E&)5(_&QK(H)#v??V~!eBYz??N0L-^+hNZ|v+={t?qKLLDn+lHjGlJP-3(f0nqQZ3D}O&I}eaOEs|eh#X=BEG-"
    "h`R0}QWw1gN15(zF(4H<K6X$c)KX096Gt?azl@h_>Z?DqqNt?Y5gnCSvaLA;uuIH45`#bgW|6zMQAGXo~R&6apBY~lqGNn*`$pu*"
    "8q!LRkT4hoo^5c}{%ENf{pW%IA4g1>bnbIwIb32CIEUH5D2nh6GPh%NfrSRTn_K`66c!iLaiaLq9AhIPf=xO{HVT;j+kXW|uPULF"
    "U6H@5cC22qEpeEXeh>uY5_=z5-"
    "PkobJO!L$uAcsm^|KljQU`<+|xHhq{qnz+?nN0>~jS3j7V>L1JQj)JC{=veTEw_$(t&#PH3&Dxty=KGCr0D(8CO}kb8)LZS9dM}v"
    "H(BTW;^J^&;3f}0p?6>Tzxw~aDo!+M%f5n}<w(U8>nQ+hsd3nj{-TyAyiYIGI;p;y<4h#x-"
    "|LYj=#&p@<`tea*)uGsUUkWYtZoE{-!hgH6Iqvyx2QdS|8{4Y=9I&re2l)sl=3?Io!qtFJ#rAv*Fyp`*+7-"
    "8sC2@9qp>JV@jApI#wddJ71`K!uYCBi`iC4h+ACIT+4J|)7<`yWvw+3byc*E+@u=r9uai3L(%9--upQtN5lWU$n{9!w2_eF1bWv$"
    "D+5BSW15-l50-"
    "h@J5pji?=36O}$XW+c^0M+GG0}>V8WS5Oa1x9Z>4(C0X>)?sbh@8d=wdJSvDLwasvF{0ck_Jj@lsEr~%h{7mq7uqwFfGlFq$v%3J"
    "|0R&<Wf6AVmVEx$OI)Cnrq&QTKo5s9cTCOxZRrZE^<#^7kloNrPMunEBbrM8=RkChT3Ie)_o)t8g)lG(Xl#>pX#sq@a_H$3knbw="
    "@1ejt*3=F^*6Jw;?{x9sV+*7@YJMFKrB#WL@naBY`7d`c%&Q1_$}AWJ`{WSR`OLCd2_lk>%kq33oI>0IIu@|HD7e6-"
    "Q!2581ugOB4ZpwgFVtV&C!AQ?<BHU;4{Mxb<6d8K~LPrBhOx5sZg-TxNg^VY1JFwW`1u|w!g>%*KicX@_S+5B!%r3B7{{Rk3F1Cm"
    "!jg;?+h7VUQXOBq?kDO@z}%Jatfy$DSRTaCDKY^T-R9N^Hwz-"
    "3GxspB}VB6h_hL3qBWAcEF_M&e6}3#=UjvVVmiKpNPsF2)*Bv1frM}J--&-"
    "AIs0QPKsfUwh8Ou$)M&z+51bpGE@3|K0LwgZGg|Ca0l5|Zhio@a?QH*b$P}W+b~9Y$Ts4=xlyJ{=&7myeJ_t~jaL-eeiO8s01}l1"
    "qvi0?M;_HO7hv;X&-$k<mn$TjW3dqnXHGPSK8ZvoKAPeQ#u*q-dO)-"
    "*fmMp+GA5({nowZe9)0K(IoqP?Oz21ZdIW>4!^1qo{!8gY3K0aIWzmYgI0l6vbXU1208Au#(8O{UlD-yM$X2X|A@3q15b*j_b>LB"
    "tiN=J^JJfPIXfpUXPG#TCtgjBgbKtYv`6FWU{=E-"
    "lw|8)Q4oyr|d9?CP2$ccfrQqCx#q1r$YskA|LHkk{<Saiq~o)e&K!k)oxY5jQ@$-Zqb-Gm7l9V=iW)||SOg$Y}!idKA>A^$;qZ8!"
    "jRFrI*BXt7gO^U|*AAgaUItw~9J2W5^NE6SMq|Ik&A2yoNz@>FFCfiXT$G!(-"
    "v4;qJE^)k5NFL3_rKPB}JA#VYfaN!d@nBimR5w6Y2+J&>QJt3`Dk%n8>lfG`8865hfweC^)+<0$+O2?B3cO%WO+Wt@2uI<*b&(|2"
    "`LPFFbHHWvjE~%1`QqAXxX_|W_H)h@?>ZwRb#XD!S)vlp%Gu>td#d<mJrk7$cN8F(S>$S)_>E*GY?LWS=PbL*pTdUQvK~o#P&HZ*"
    "<9;ZuVF~`#0QZ<|1e4v|yZ7}XkF!J1Cl>k60z>OTU2tFpJJjmkFps+Cn9|_*V&0H@vez}(^xwuW%F0#RO4hpG&w_I=O%aq|_n+5d"
    "~&H<nlV4ERy?E21Rnj#tFae+aefgly<<$TAF>8t2l@%2lyM_l+)-0<(1VaG*^iK=}|_r1Er@UGiSJ(P>A4Ii-keCWB0-"
    "U5kK_*Zg02-g#p2t6{~|7!)jL{q_#in`64dRI0#h<x=YnXH~+TH`$!Tb>U!1_Y!cTyuSq{-"
    "?_Er_8P5Z`5Bi|D&sHh)6}h?i&4bLr1?HLK_ql?DO<Se#`;RH&Doc_{uO*RwYy=;<5iMlq*D1)HI#!fKF2voO?mHfI)d?6r=*(%H"
    "eGBR+JWB8nW6O+dQ4U2saBI(;|Wo<!I5U=J6RhQfa+&-"
    "H`S{xb`oCix?u>L@ehZV!8^*Swxe>LbS<x3*o~6j5aG(cUESbllO9YNbHXOaS@!DVA1AeCC8I;p>#k$$G1%E>+V7xkQg4?G_2m_("
    "b(lOo-5;gY~EMhVX}I%(C0+QNSdZPbVK}U%F*u$688cVQcC;ddJSIB`8&f4NoY7X1AOYu-!1etbz|Cq1nn=x-"
    "TTty?;Qeug2Z@2M@fn)ptkuJ0{vg*J4Va&hS0FIE9x*wQPqU%CXIXo-"
    "m70fPzPF*LUDtITlIUv;3<|aV7Ch(#cT^~0^%p}D}OO~Xg=Y<?_)91G(e@qSYzq`bZpTeY8i3})t)io$88yfWChI(s=bA#0dLGz7"
    "5D{bv@mDV!5r6{EW_;_Ghr)UtWS^CyCKme8A+5n=sG2{;FDt1L%aWh7!JOlX!pN2#@FmW!%B*^+W&FqbaCgzOko{3Gnpky!Ihb$s"
    "a}DaelXnvOro+BUxR+qw=8Qx<sO#i@>Im6=xcg0_IzntEVmd+Ur&NUYpA&hH`HUJV6jBPZBL^Uay2?&JVDU!hya?jkh@flYiEO@c"
    "79_9t6<OOzStp%i-~4Xm?WcjJMCvmG>YY+%tB6z%oX6r(bl<#>p3Lb;s-"
    "RqVO~+dtJR_ks<~uFE>>u^IjN9I@mGU??owNl(LHk%ccea?7@d0fV|^>4lON&Ev>XlqP@^KGtoum(jOp;W$v31})>AWBQnWRJ7t>"
    "735?re(O>YyRxe>E4BnpDj;Sbw<BZ8GbpivU)P8DBV7?LmI;R^rrHA6Fk&w!T{bIY)mh)l06R22xKrdq8kJYnH4e$WY9e~v|!WI#"
    "%aFgTry=1iQ-"
    "W=dDqjSeDPeKTfK1g`oz)|uQd<xwv3z(jhAqv@U5f)~1Ebk+0Pt?aBvSH8v2c7J0Mi^#DU&3|J#mNhh!8=bF&6w3{$N0oaxeJ_6O"
    "Y~E6T?l~O~hVr$E<T*`2Op3mym(sk&0v%&9o@PtYEe2@t_lqpI<QwI1Hw>f;X-^Y>mCjZKrOdM(nvQcdrpcwhDB0bf$m+f~r7-"
    "V{cUv$hV7`V0SAvOV8WXPnko`(;%}CB^8}8J|*HjF5fm1iiNLAI<h4NI~$tOt{$|qyiFUE#)27zR#SFC3H1_7mK8{4<(OFo5=uI8"
    "}H{-Ac0rgl*&0vL7+Gx#(EB1u9mpV8-"
    "$+DM%mnky>nLAOZL&lnhz(YKB85cAlhH`!@oZv4Zb?RSCEKb~Za8%`^?n3~EJt3Md(b{lfzveRJ&tmg(KB;(%9Va_%$c-"
    "vJalf9%yZ(~tbpQ!*BNez@WShH30(tu}wS@4`kNF=6SJn}D<Rvlh<>Ad|yr^77CKCK1TacMrLAgPp7KF>gnUv@mSqhPEgr+c`G%Q"
    "wtaE@}*O`Lbg5CAXCFG||sNUWhNw_;;?q(+rM!oq#Kuzh(vGlQTi-T-"
    "2A+CS*JPEB0%=H4*tT$&QRfY2+B_;?l?^W5_ndc~=~j&cuI@#TRj*-"
    "P$ndRW>?9ujk?v6!fBP%OSrKZ_^XiTB5gzj{BlF#_w1W0|t75woT`HC!op)`i)qT<6bfA{Izsv>^IwE20<^_wvo<n6t-"
    "irqdXSkyGw%>;tm_^bb?X1o^-"
    "$RTsg5|qDMmA{;His5<33%IX@3auFdgVe~kr|2VMU|FtgCm3%@PLENAe&ke3>t79T>|_ItDIA5Jg4wuxM?l`GSsNcimPgK)MwtP%"
    "#=>rI^kf?lv~Bcfl3_3}rRB=}kP3<?)-zc;iA0l6^S21WBLJQaJfccK{dLS!({4DtSE9Nyh)o$k>4%-"
    "mEw^rCM$nntyN1M`rR2W%vUocIj40RX)ySDjvqlj+1%{msbdw)%fF#y+23kXKw492erKF-as;SiAB#C`vjL13@p?u9_w3UdB%q$*"
    "k{2;Cafn83esx*BtXD(R2L9m@?U7l{-"
    "N}Ns~AbdI7KJkZin~<@xyy_d<(&AYIBAGYr&#ZDC{pi8Sg<wZxu9+6?WgLx*wTmy|R7B(`mdu_!`Qv(EZrG|HVJ8$9hK2>R0S?Ku"
    "EHiGY0+B)wMU_BX2Bel!m<<1Oq8y#y>)U0KYvx}hKs*)o{H*=qBPrP>$Or8LgMzm$J3Ia_h$EcWI-"
    "k46t(e^<HDp6o$8Vzpx5p%u#WEw7UPGYmDDQQki%bDT(Ht@oPbBV3@jI&8a*W=LZeaIBAe>x*hL>v0!<FX#nwe|0+DqM)-"
    "3$6z3rSGU9u3f9VyA%gOF{c%q8w67{lFVvKVLwOuX_!Hjkgj}wLTi`7ul)7F9J)Ub|vkPWPXk#KA%b0=~OnQ95Xvk+ZvFTZEI=1v"
    "QzBzm+_os=nU3fbEBIt%piV2?KOA`gbH@Pl)H$ekzkhMd>5X&@c@|8s(t62A<I<6;@cJ&>|HP$&hi-"
    "J2$>DrqT*Mb}8lT)aL#W$!_iY_Q$*_(<Vv6y+4N+k*9J?ktVMlwvK6gp!dA6oubAH~<<o92Iq?;d?mvr}0cc1Kt?H3_WrcmtcN9G"
    "N4$l^q}=|9oYzEqxxMXPh!TmWZjq>9Gmq6yV-igkpPBk`5#wYk>JzN>oWaGQ0qhjPXf9OOLz;RFks39br-"
    "Z$n<*^Zs7u;v*g+ixG=KH&?^Yx6?>Cj^5P9+OVzr(?AYc1UfiN`AAS1E)gP6Ie0<gW&4?Q?hnxRFw}2-"
    "9EKCU=meY7}`d6%%cxxJR_V~N;5%u`bV=l18zf}BGt*vt$;IaDRD0v*_m^?SmB_NOs^NQWb@j_K7<Sq6|ES%we_j~<4A=b+0m0t~"
    "mzHfI&3np?j<Z1ku^F;QN2Nyo2?RjCX|1%imqFu`oZ7CUiC#0hLX^~&+uy2Zty%`ia9_m&HhUxH}YdPFtKQZg)k$1-"
    "h69#gjZaDRg9;q_PM3WEljc~7kzR?%}azSod*Ogh5NO&(Dnr68NBhRZy0s^%#LrD$h7t~8bDp#!Rn+S_xm+6UI<kUgUt94qf5Q&{"
    "hFKa-O0{yf)QQP(_@vrTf4$^C^@@R}`;F|?MIcBTilNmGjY1k$*65h~3sS0j<)Is-FEXSD$N_CJUVStWFx|szimBF{_c2*h;Ua5)"
    "-`Ff9KP!ZaZwIOTw(7Y(K(32w-"
    "NPKY9R4R(+vwS&MsL2G5g$P8tJkk{lsA&a<;=fm7VLJHb>RR8kZn5oguX1mUrqkJgmK=9r;|gw2A77Z-q(o)es-"
    "t>gB2tl3b121Y$3k^;rgeh!cRE4|w9#pg3J@@4!AdYuO=H68Uvat0Ta%Gf!{ElsH1!gLU7(DiY%EpnzO=8!UTROelV}^kZe{(s-"
    "&waN1QY{p(=tePro73bysUPZ+UBR?2cBgkW*8I$zG@%i{E43r_?t^g?#`8))1Sp#^vd8293E*0<I3Y;$hKMUEB58svp;YSkQ6zp<"
    "#;G|)h=_-"
    "^jw7hH}<S?NCiVN@^!mW_L<tnO;n4`p*m!1qwYGxa8+6xvJvy{mTdt=kqT$^!}OBh$PG!NuozPC<+xpp62bt6V({%l5TaY|RD3O;"
    "ez}wXcM*qAi5=0FFamhYM&T)Ez<i55?dT}d(sCUy{iA4~mUrdr=}}i3GQHW|W)UrgOqw=>BFKYvo<jKW|5eF8q<P^-"
    "@uGa8Z=CLo+d{C?XT4s>`xv4uOcZHYcA6+BRU9V*=FjtuD{~lwuQ<xCn)K&?z+5UQiab<vSb3|K9)5s56zjD7ue>+^oYKd%HDp~5"
    "Z$T5;oj@;;IViW;YI49s29?u={X|V^`ILOgJPZkB3n9%6uM8|T;QBrjK@z8iG#VF{k57e3aJg)>z!h9`N5F)LOCCe9$++T(39{G("
    "lEuo?;yVL*j~JR7v<{LYUqB{X^XV>OQYAfe^$n#=ziotD%C#XD6=fE5ii`p&VUNVI^o8=6zlt^)e4~CWinigu^WD7te0+o24L7LV"
    "Gk+{DF;MRL7zn?~l~KGAfONGNsw#bI5MMs#yPzj%w3h~~6q#EE({rhtdoKB>sk%f{0+zKn6){B?;S{v5#YLD>G+Ukaj!Ms9QOxRS"
    "DRNhb-xi(E?eS9z0&`{ijgBe_pxY#+UI;~^8{lTo{EsIFv*!m`1keVo6q#!RH(<IJOmQ`1mZP!?)d+oKAp(|}CFN5rKs8EFrwg2@"
    "bJ?!pIpq@(cv5!UX2^yj7+_Oma~YV;V<iSZON?g=sody{M_*xMQin{D#wz^&)=6@QZZk_=n+MfBeJTDsQGa`=A2c`FcF?QbVIwq6"
    "d<M7_`CEk?WC|cqMl{{DtV(96-"
    "fFjY3oKK;O3;@9)j$1KO?y2^DiC{>djXagrH+##Weu=$nDktZ9Jft!ENKXiq{OA8rN~_qe3v>Kh<<KJ2}FC<hpfGgusUtRPLWl(*"
    "6ABH=PpdGRjzK(6~_kZXtH@?^1A;X>DebIqPvn7UBmCB!oGC?b2dGjA{n(fZ~3(}O-"
    "F(>vQ>9GUSn9fLeO(eQ!pS6sDrt>$Z_)T>?Ql0`ur#o__k8l5g?7FrPp}&nxBp;`y^MnxyPQhE00*8`E@f$1%foNHEVlbl)=4J3z"
    "<pM#PiOZfs!@(ZOBW<gEYi-*Lzo9Qu40d3<|nuX8@3fvSDwkeq)}@x6ui9<YYe2{nKwUNA2MVqpc_Ry|T3pm#-"
    "fTkNa=<{@C+_ZpMH#ptf#WrLpAS*_V5mq$Koi$Mfsni~?s!lZ~+%?PF63>vUFW_+xL95#AWGr+gh)L=RaiM!)fg>J(c^w#a4>a-"
    "iPQNAV2I2X!op-"
    "~RW)qXU1@JqWjF!9v(^0!)M{AR*?pQ2PsWb3>`*A+}P94i)FJ2Bn>Jf1#l2F`)3~ff7l&ZlX{SL}|ALD15XY7O-"
    "d>37jf`M2oZ6Z}`Pk;w-#|!;&|F5-BIp(!)z-"
    "h9uuX2gxi@xG#7x#v*yj!i|)%D)3WuI8B)+o~11$P2nkEkrK25Opr^)8W&46Qrkc(!P#Z+tK(oocFl<jH`^gu0Wzv|%t)!KfeCUQ"
    "Dr#87xHL*{Q2IUB@-aGl8SHgtz)2?4K62Gb#ee0M_~8$R7WlRBfONML;A?|7!$(TrEz%8BGdZUc;@vQ0F7!N2{8c(HvyOGzl|?{-"
    "U7C1nVH!H*3)rfoDDF}`;Jc#HAsU1df4`4=oj}Yz@h_7zV@ZT}H7&fB-"
    "?v=UtplD@L}!b9B&eHVEl^fRoQXW5ySW<Gte5?r__5Q(R*IL@cmw|Zpj((C&jOGPDGrlFewMvd>J<E=CW-"
    "G3yOoE2m1Zd@$OPFoJ!B&4MK9E5?%WjTKyE_2+nz(6iiJ$HZ8E<f)tw~Wq+dyY8r4Fmy=c3YO*VSUjdr#@&!e!zBEvvE2LnF|oM("
    "E)Pz1g|^zE{riiAw88+PTj?2tVZ2EekPeUco(+Utl8ZqE{IGZZqxZo1ysCNH+lB8>7)5XgkLWuKoib>sY2HTnN6zH_RJc9e7BcGv"
    "GnIR{Ww;-"
    "5!{7^Rzj78kiEzW$p`8)(*MU8&9)w+D+5yJf;GninY5+bWO&==E>a&%_cd^L@+V)?OWW5tMq&c&R*IK!%@JZ`ILrBAd6x&rb}Q>k"
    "U>}8}}<~0T#D_Ox$E>^iU+4h{y4z*dHbH8*AHREbZo5pve#jUqs5c8^KmXBED^b60tgN-chhec$W_)Y0Z~Jz{MKXYy*^8r<{p$pj"
    "Mur#Xk}*I~Otf?sGL*7mJD|3o{x1mcdHDx5%O{tU6qt6f2ZLXWq&KiYJkLi2}-"
    "$b8a`zujNoyCJj=vK`=sheuyI`L!F1&EANz4f6s7ccrZrOIF*W-"
    "40$!+nE)n5f>I*jngc*jngArm)%Gni=ad}R8KpVbu&COzP?KS>0iP+8HN@n+&do8#ah)@?WH^HxlAE*gys%RN`LZ;ipx1CAV&tK`"
    "oqtXdp>BmIyY4~>w16hUWX(|?EXI&>r}N({w23EcqthLI#ln9$D8foJ6;LAZwY0$7{|W2G-"
    "8%X?wZaZGMB9pazy%71OG1lHI6Q3)kuLJ5ChzSn&suF02wG9LDS=<9NgAK!JPpwpE6cZMn&yLzDqBO=V}6xR7C^M3ziJi73!{>`%"
    "lt{-13#}k3S-{)EP*${pcU$xYpAC))~1P%lVr^KfH>~!Ur5x=h-"
    "ihsZuhl36MK#mh4;+h$+giL27Ufs9eTdwv*6GQe8b7O@j#$^BX+`n_~rUL&1!2@x#QQNH5~@6Xg3{gGWnzY-"
    "Z(aWtK(b!Zooh*$SwPZ{Ybz(iUur3j+6<5vh{%LQ`yeYv(>8&5vlO!@-"
    "3C)ojwZSr{M!FHp+5wmAB!!{JxeSlHO$9uh~N6!+B6IggPimaV?kS>4o^_|ClD3;%k>2=O_l{EXI3Q$3~o)Y7S=P+9E7M`JUxyZW"
    "7LwtOmuGy>56o?pM|YW#etO5XPJ^6X$T{CJxU>;wJt{#O6BuI=Mwy4>!MSo$+E#o((XGv$dM<iIbz$riV1s(1<7tB~FZDZp4Bfv{"
    "b0H8EY<qRY2aWH~#nHf1mY??ytfv;y2v|maEJFNo^dP%kYKa6T@t2UKoCeW~{UsEiK}zP{dvRbuNu4EggOz=Wg5<8O`uihNiHe0-"
    "6?uD`1Y&=gMd{nj$K{sio<-;Fk|R-9ghkk%gQVm8)P>epEkpsUHzGI{f*=BEmO7r9~O^Xzx??>!*q7-cOK-"
    "5L3X?;tjs(lN$z}nszKt#5h)Ta_t#hTFf=TWTZDpC1cT&mL}&6)$Hk7e&Gb(!!b~XYLy16;WI~l644+a>V-"
    "_GmQUB30T${|6cKJr+A`S5Ag>k!%Q2xGqKy_@!sP~<CM;=}(ukj+2CB^O=q6?&#Wr*e35n@|Y0<X|<=_;xoKM8D>P&Ci-PmTm$|K"
    "euVKHtp!KKAq4R{^3%W!cmo^<uPP(xXi{TRQ65A+S_2787l>taAki!q?SPv^c46s%Ei0BJShQ=GqzHsRO%$+|2;OS?E;TI{uMKUy"
    "bmUt#e>cFNJwR?pSuxj$`i(usH)qARSqGX^6O?Rs{!&A*&^F}IHKjhZ=Af<0^F6fg^K;UpLQYL4!QGgXW~=K4X++d-"
    "t_c{OvRGaeMy3}JvoA?lVEPQKP1i_ZAZT*@9*fj+24_cvS)5xood^cmqHPB!YW9X|WhbBjbiloWWeHFqA0>pPzF>iMic2U~213o-"
    "6@{XPV#NGOEaxv}#^Jk{sQ{%?{z!iT#W?)O?uSEm6Eg>cuMR-"
    "4yqIOTa_NeRzx+wB6Dk{AREk#(5qA_nSmXQ+KcK_(ji+xI=q#Q=grc$@YyBl1P_LOGnj*NpkX_uyJ*JMcYM#DIcAd|OVp??ik@_y"
    "}I_C5@qal?Ux!0XMJ#27wqSTQ8IEy1z!<Il>Ubu&a{_>vcZL-"
    "E*IDDi{ipwwxnHhF1J>UtC#;@B3Jo?*(E24jb}<oX;qmflvss!)fuAs^Y&<x6q01NA^$b|LqQ0<qjLNe%JR=F%=JicsJ@`F78Kjd"
    "5Tl<wO)l5X%qHM&`0CqZpKW3w+1|KZ>O7!q>hzSw~a=Vm4%uDc}-tB<cXXiYZ4*%f*w{0@(eWv@>*a|-"
    "utO@eSFqMAUyWbJZdqR1Q==p$y5)b^(8vezR(38VmHSJt97tQJj{qGP__)eBEHqn>S7_>2=J|fZ^@kJp_+ch&?A27LP^FAcCXZL{"
    "XQ{ni}eSr$Gt5EmGw4~#!>qPP+)Tfj7`@h4i@srW_y}&Lr1a`n<}q~W!wnO1SnLK@3&%&`7~dVj!$Ol@V0Cuo5mEh6zHl$u80m*8"
    "AR(L!SW5%_|}QY^i9N-"
    "?1^o3qUgTH`a2z=3%=JKpqYJUMoob^pcIH>VHw4$NV5VB)w26qDcMcVA|Xq7j71Dc3Gf8ioGZ0&hblLBZD@#Ry)mZFY8@^GzLsIz"
    "ka&*n2ZOp^)Lgo_T9D?jQedhI`L2SHTq1g%l}MWjHH!6yf@WOHDOu$+RAcTt5s3>!gW&x!tK7p>3ytBVKv50&K{GW@KbGzVxEx?&"
    "SSc`7hu!>N{g?P6p`oliK(_av14@CW2AJpZ4TSOB*0|DNfc-IX)jZ&Q)qGte3-"
    "ha{+)@HKS=FhTb73TrY%^@Q_0jSO91F#!lxtp#FMh9k$8B`lqo*DAWw^k3@2Sm%ln8tyCGc+litP$-"
    "O+vm#TSqpg9k)}+#WmfQik+%`tmlz(8adMDcm`qAS=;**&)DZvi?3~Ls?-|niyQ1xN=IT*e!!K{o;wZA7^p>h)jiVun-jUKPXv37"
    "1w{or4ue{}Yxb^;&*G0E(4R%lNW#4mU&MPe9`kp9xD&lI9(r%huz;dQgKb{wkMhdY*6%p3Ip|ilI_({=AvXY^7UPC}X;Hlg#*&L_"
    "bLskCSqvzs#k%R*xj0p*L_;48&5K-8BL-9y(!FQ^L@oHX*4*hVw|6S9a-q4d^(*V1U!Z0T#3F5N&{er%UQ(0deS-)6Mo|-"
    "h<0z=b+TjGc8#PBHMXy3dxz-7MA24+U)Ph{S$q36=V!iy6QBM+{FZ(--9$py(xFAZP#vnFwjTt=__IKZkl-"
    ";S82^&|of_8Ty%;>PlbCPG6sO;sEjU}hazO%K?crm2txX6=Ib0<Yl<WJQ2xII(Ca&3D-"
    "(_J%RQs)b*d3vJeOurTKrOSliyTjW3$|jOYg(+}}Z5)7lr}++5b|hblg5rDeGY1-dLGLTZo}tmonE?|c5UREMjS?-"
    "%&DX!n)17vzXO+inGzyRO42^o#44N8=a3h!Zl{-Mfll@MB$dBEClp1GMFduy>nfWMwUeD-M)}N0V8V`LPDK*aECj7h<-"
    ";_m+ltG1eJ7Emt!wNC<fu<s-M&>H4X4RhSLnT`!ofHIW!j67RYT^^BaPhXoBMUcRrN$bN_V{<IALUY<pWdrQ=s@ep-"
    "r*%AzH7lE8P5VvjZ1jDI!o~tn@=6xDlOEQ<s;Qo8`~1-8_?-C2G1tcgqa$LEihk;3)La9mvwsIWM6X<aUI|-=!tQ9i-"
    "vDKL40HR*1|;UrPFo+j%l`)^016$#!%b4F<n)#4UMZ%70ZsM6pB`;UhiqY;I2Avog{i2@%V$9(Nd#s71G`{yTFdf)pg+&_Z#yClz"
    "nmOPC>yJ@irEyp6^`3FI+wP7?fZ)rmq3~W;@nx)%ZJYqk}{kk-$lfF1RLe*Y>i79okCKdj{7Bp1&D+QOuR63%0T3IB;2-"
    "AH7OVgFIpq+pt@}*kv;qJ@AdCw61*1@7u2G)*;V1f7urMNMAVzw!kUO)VZ)Vf$z&PNYh>?eW%^8$bKuIttbWzbOO9$A9;8oa?U^S"
    "i{r13c#Jp4Y``7id<WT#gHD`R?ZfzfiqF}Q;>!bZ-SMr_J=P!a=^&fz+1PUfz8w}F3hYysqcqb|s@~-"
    "ezu5!tLnl*k&<S<jaiF#E12=lWwhNhHtv;VlWE)PitN%DRuJlB=XB~C;pveNpSL-"
    "O~MA_Er{WhA(5}oJceVH5l)ath_(*%G{bX!gx##8^R^Q7u<6fP0l?Lu0Nbr@vA>@aD0`kvwa9-"
    "k6#y0s1~>P4pjkcn;sZfu`$FGw4ebB|{(8J^mX0W2NPmI2G)`~ptQ(^nph`*`NAn8DKFtP0ka@J3wyM2H)YbtB6=LA$rp86uiVWq"
    "?bEH88;f*WK*T3I_b8tSi~a;XdHOnBmc~r(mYT-YQtj-SqwbKg_eb<u(INA3`c(I^5O3hK!v>ivH{0iW~WRUB13HWIL57SoS8Vfa"
    "y?IhkxuY#Y;r2L|UOPeDc2PVtnp&$aMH?KuS@Zj=zY6C}hvcKo%WtIUy8;QBEEKnR3LMop$vkbkvv?;OooUC6=#-"
    "v*4K3M7|Z@ktOcLppWn*&}Kb!Z?yzyIt(s@9l-"
    "a+=<)+Ks`5lN!|3|h_J}vP^0pl;XYnlLbeOCHPxsyx@BW0LSmaCjHA7_VU;s>qw-v~Vj_(XwOa7|6k7U(Mz@)>|D)5ZdjN&-"
    "10jbog=uj3L@5FCih*!ZxO%#1&*qGcOvu<T=$i57ESoV~e$muXy2NR@ErTIUY=6#1x-"
    "QuL^Iy5@;G$6B)ACwIFjh=^m%BKw@N<+qw(qRiI6Qbs7OzXu)_mfRwK!nKuBnvhj4qITlgq!M7u}BrlB3x}P0otKBO1bekpwQKPx"
    "0d1SIj_W9<$o$w@F&%elh6(ERM_1PuM}`&+Ah8_kL60}uv~>&aHB9td6o216C-H(iCAB72i&iAr?SO*0haaF2A>Y2K;jkG-Xs+KQ"
    "Ep;_^4OR^6u+9=HvDWTpNkXy-40jzb@3hP&FJY6T7wj7eiSM?{h&Q3*9`ZpQIjSj#ikiG9SZBf#fClIk)y}U4C~Lw$Z8CGW*-"
    "i`L^Z%10^EC#7G@1VyfF#Te<M$~<H8c3^jjg}3^?+bJS>bb+I|Dej&>Upf%1Ftw@0abtg*pUZYNui?M4=;1Odxw1UUUG_G`Q~5jj"
    "mb+!#@#&gMZEXvR@4hAh?`(e#=#v2ce+k{6oea<}p%@akP)0zfayEA}SNeKi&HTzrw3Y_!qY4!XV@ISnA_1$)(TVttgE{wTf>O?6"
    "s!fae6**KAbC`1%wy^ulj5=$MwlPedtaVye?WWIex4%VrStf?anHwwP#Tm>?7GFq!8bd@~Aq(Qa6y9qkc*5P<(A{N}<t?6Vg0Xkr"
    "L2)Ac<EJ{=If@Hg!)S09yc@ww7MnRbwy;SJyGY%@7fz_a0179@Hc*jzu31-"
    "uh+>s{!g@bI1=1pUfZhu5fmhTVXIT&Qi%A5k;HR1h8BXWwfX@kqe!`*dVLK`+|NR%^FeX;f|+cGW9QPL%jrWQ0x=CEjo^!_D}&y3"
    "Di3n=?Lo<k$=(=bEc{U)-"
    "5It{b|+LI7S_NCiSK<}22iC&kCI%SPMKE7|l@Uh451{7gjTD5xdYq)WBZeo8S;p9)tG<XTX<`RP}T5gD>)eD2baUtz_59{y7Npqa"
    "u;f%hn?iSjq|2j!3wwPDR)UmC}Lu6kj&#=o%buT1yGKU8b0jW3p}Uo3MZ$$R7<s!L0&(cg=Er4FRgK2zn-"
    "6h+!2IbscoTEdK$+=E_kg119QkJ3AlI^8!ADc*;hY>04fWeRqBw8B*lRkOcn(h}D)NjMX1bcU?Gz)=fkjO18@@&uetcKK2L?Wg1{"
    "noYLNkPV`V!KKIEGVmyxbNTCut<i%3*({nFP<o_QfgicuXpFA*&J&N^8oz=rvWG4Pm>z#C@cS3*F&^Enh4y*Bu?pV-ann2*x?YH~"
    "CDa-<_y_hsl*-"
    "R=&iFOPQ$`FhJqqEZ*^X5NEDG&1lpCO21`ky@4wQKPFuqd_&ljVRKW@f^;rTP;rAJ>AY;w9|wG_($b@F9ZDC@!L);+wg4~A0)wqt"
    "z+lN?`A-"
    "h(&N=3ucRFOP_)xeUMaJXA%XQ$5EfHlJIpS9!wl?LT$c^ti0Tb1BbQ7fBX{A0avpXK?AUw*qEJ@{_=yXvy7V{b7K}h9m})9$%}#n"
    "Qf~i7*~FqwssiKtWXuY#ZSt$Na`mpv|VVovff3sW@f>rM`s=E04QZjQy+Kz3^-)mB&Ln$1ZcoZkG=-"
    "%TFSfB!u9E;p^>GFYZ;wGN{_WB_!}_01Bo+tulp62_aEg>B`8P3b8-"
    "5WB9n<Wudi2G!(yte!z9NSkX<4F+xJSRCv#E1kFoaOh740}!GM(>bskbIi54nxCVzymsz<f63d(`dxs@D?mt+}w&Y{c(RknWSaFm"
    "ecKW3Pw8l?iJM<Gyp<w*S2PvV=Cq<w|@fy8sU&iW%hS0J*~SjX~&=DQwxw0ls(yb%wDA%}fujIazbLragmI@qB?ND`uXqc8OX;R7"
    "~8Q$aFgrAHl121cB(qB+gcSg%Nyr)Sc_qajaeqnSr$0Zxz5Cin~&{{1uceOsn<tHVO%tTLA|UHGIIhYO&`Gh9>C4Gj1eMroN`-"
    ">QxViI)PNyl=z*{XX4KWxUhjZNkM^BwAo4*r=wl;rb6*FZ9-"
    "Y<n&2&CT41gJ%e4KQ(_5;s@jPXUFDGnziduEQeV6g_x7uJ>qF4)SHwubbtaDpBi|f3A*9I8N)8W~S4#cAm*_;f+wpCPmI{Po)K{("
    "MmeI}iR-Il>3}-Ak6?2<c-aNZprlX-"
    "4{hF)W<=gm7yCZ%`Ng7^m2%nZgF_xF9kSG#Sw=O=>JUZz}a}eoYGz$1Iy}O<Ev##g$CmjpL*f;FXfO1Q6ERNaxb6HDGvMA->-"
    "HOQ9d3FZOf<%!82Y*JXV|XP1PIR$oc;k<IzJsnKpcrVI3FCXNKH#_1Yr`%6n~H35@EK<V0D^J0IWWf3=(BXJD3bgxnf>P(WY#Ure"
    "$D`hA_s}1@kFJV_R@EnhbsH_hY)gdo4Lg^gzZ7)VRsM=(+TPGzy9ZJT2j$i_n$h(1(K8DLRPDNv9uiiFU^H)xf=bwWUK1VR&1>6v"
    "6`0kMt!X}^0fI>(`~l1!+m=fK%4{IGB8*Ns)s5wC{;gBxK*t)B<*zuqy*Q=dL2v~`%($uGPe<1ydO*LC9p`Gv!GKX6-"
    "cJ8jkDJZqeCV#UN}fJK&41r6MSei7u@U`=6&<=ZfkfR!c@c*S!{u6p!Jk{N(lgq4B7yfAb(I!j?*psgkOj85q#bo<9Py2XerVMZR"
    "IFmEbpe(<#@->Bi3OfWIeAsVu~zQ!JLjhaJ}hu!e2f-"
    "&98MB8h@i0N{WQ7fbU&Yy&<u2;nfao+m<=76xmyadffPL9SMcEX$tgKz_*=4EQfeAT#Ed`<zLRDhI^szi05Cnxj6tslYh~XQsfON"
    "|FSFK%M{~BWs`UO%k`eV>0IlKc}6tcvTStPSoWS4^b~n*0Ncf#8V-"
    "P>r|StC%M}Awip({^PVc*_=i+lBc_wEX;b~yjp;9ES1>ch8T%ASNo>h)AE|iO-ld=H?hvle1K-o23s$9mYa)H_v-"
    "#pPYK;dH;K7&CYjb~gcYKml5!N_#babaPw1cEOptH#$NBP=x!FT}4;v^3~gIBf<!G|smMyc7vs1ztv{<1KeZr8KdOigBChmQe$0i"
    "e%QHCZ*&Pgd>@XQYz&*l#k;LyDU&pTFQ0|AcseO<oZGU<V5|hDGjjE;UkWvsv1cyAPYD}K0PRECgOB`uD?EcVd?>E_wfm8Vt^@<*"
    "a8-0ICm7{Qf8E+^$^M;a#w8bC#j`IaUQ`_YfQmQk-"
    "%m6k$={<YIL*xsP|uk8%N4h+Li7axi%al^NgGUrAXTfl;2>6xx~y;0*RMk8EhtioFZq{y2u$SSy&WaR=6sQGprPutAVBAF4TtnFw"
    "sKJ$MtQoHlo2dI#`O-)xkEKGoRX-DO+-"
    "M7`NEv&)+u;&8{;GI7LdKhj2(6#RDNX+E3`aH`%i`GJSUfF+~>Pd>?P<(y=g4!17sm0Gd1A$UJb7RF=!CZpG2#aVnX9`^(k8>6w>"
    "x*6Rd=$_5+pTB3hvM+2Bd<X6%npY_WQmvrlR=k##Q41DC<7zbORb0cXYEFO(KZSa%W4Ic~h<%xJ`b+hRHPDgIxxsk9>%a&9$bOLU"
    "(U7X^3B_99NY01&F*xco%v`xP@Iaz?{gniY~aE@<Cuk{s%%L=x}{X$&k<^a%%Zq3ovl7Hvi(A)OB^1$=TVL(79ymeQvxMNjhIZDV"
    "3`SE;Ej9**~1)V6{d=Zbu-t@J|E1Zbk>i?3**7-"
    "chcr@y;e)tg_q|a!lq9GIT=8~5oUfgKPFvN@Ppd8{Ap`}CGGQ3TM%*};A{8wFLD_#w2Vn)Td-"
    "NR=Yp(Ce5V-<2R#WW03fP5l;>}5ilFbppv+(kVIFJ+OVDYDSgA#eqXz-w_C{D*gi<}IE*e8jbSo$dlNhbDY<Xadr*zfcWbduo`R-"
    "r&3~`ts0pbeRFu;jIQHb^bKhvj0ijmTzOJn3;gmA*&A6FB$#T$uHNxRn4UF+Sot#Wz!vwrpXkzbcky}R!|&|z>6%KP(B|+iOSE#*"
    "T-VAmp(a?bA)C|)Pd4r3P?GAW;*Fan`Uj3D357kdyj%U3v@D6wxC@Z-zkHYJwpz5%+UEV#_-"
    "ai4LD1Zoi|6t7{vs83%b{TPf6zWUgQ4YIqM_q_K$(3!`TXy@eTj23vud^DcS27XZOe9et)C0J;vg8V}eVEwN)@7fBq^t7Fj?2ayp"
    "bYJiodGi|;;h849dI{7n4TiEyishUkBjRW7YKI^^lehwK?2J&s5xo&}!{qd)@jKPolm$$U~!7l`-"
    "!TqC}X<fNX8nGSnRSeulo2^T*=y*KHZcpCFQ7So$7*mO9&McQQwJ%_~jZWu)u#-"
    "1krDxIq6<4dSHUUk%XUy4;kN^8#4Dd~ki881A^jYr`MD6x9idS`)35U`pdx3~W*_G`Q~5jn-"
    "*Zj3}K@EGU<Md0OP$UzmJP*L!0@EgOtERhtxQ-"
    "Cr&3lTjER;-fRnAOd(%nKyEc{akcQ{Y)yn*xMf%v;N_9nWdWcG2{VDA{htOOL)Pq)Eg<GMN-"
    "UMwqN24pNdv=y?Lm+(rW0iAsy=<>X%FK0{+tZoo>9x>Ya6a@j90jLveD-*PoDN0u-Ue-__7Q$tMt-"
    "7r9BjeXS~iRmk&4jbU{XvsoPkJY*#w?MN=AWDZolC(7daz=APz?rGSsjgdG54wv;8Kr`w$5qp7&xHN_jH+pZ^5GgJQ+tC%;Q}7NG"
    "UcUeH%{U1rTV2tP=mqqo-j>aL@syF=8jr(=Oy*ua=PSO6Ljm;ZaP!(1n(66LUX(&Vy^1g!p}h-"
    "3D`dyS~WYo^SN6{Cx09W;mF%Ez3;_(BKrm|^ey?XxH{%*<~8C50Cd9*c>H}A@1y)}^c}5bC|b83B55140Vm8X^bz=o?|k__cURo^"
    "hSv0TVYdD?L)V*bfKHg$2AJAV9%ByFZf!4-s~WOrA+dw!42{yz44E*SO|a>f_SYho>6RsWe}cMR-uRrh$S%&Z9s=ra=W98poa@)m"
    "QhiD{I_=TZ4*OE`>GX4?P6w2rqMky9!#`oYxLZd*-"
    "{IMXhG_3O54gDVa#?7p+G*KUBX3Q)cQKvdN#H&71`xCYZ9533O!}7fAEWoR8dkosEyvCIJ0rVdqs55LHS?9&oAPhQ!hW~IzVMrxi"
    "+?h$pxVqbmEnCPhTlt=^FLvpZ)O7mS|P5xizg$`n*+mC!%<PX$vF(P;@rq5Qq3G7S)fKWV=g4>a5pVGk%`Fq?8i&Y`vy0hHPKMH7"
    "7f<*xi7Imq80zd(RZSbyL;)o!gJy0JV~AQC$@(V#=WOv;@`vWAQ+}I(&vBu&)K}Bg0k)%%}^J@(Z8_muS_KUL$y|`eX+D0{_h`5W"
    "c@>RX}KExy&w(BwrDzrW+MKWd@Db|GsVh)mDGa<H~O};^NTUl2X0z)DBc=x#Oh@5z}X6~0GSI;0x~TEm%*RE%xyNdXOHTB9I69mK"
    "A90sk4>gk?(#{~Seh#gaA`5O0=rAe&K;oIq)<&V09qNP@U@PKt|QW1)qtlE+fx0w<`C1Z2|#Ij-bEdg{Kxjyi~dslfGoCtA2)pqH"
    "zjvNF27=ZZW)Z_EoZ+3TI4ps$o(j4#Ft{zeQZd0-"
    "5+zmA~8D>3z=rfv<QUjh{>)bEllk}x6A~m)5)v%SlFPDQs@mDj!iCa=aT6c0_+PuSXtl|@BNQr&-"
    "TY(uD^>j#gqPHw!^2o4EbyqR)m33KC}tG8fglu8~(e_;kY%KIU8bb%Bm9j7_O$k2ARKGmL2?(uEgOfW%*nicjZiL&lX7r5G7(^#V"
    "V{2<MyIWmnco@`H>yowaD6@)fFZfCE{V#wUsP=zkWxy`t5jLz|x_Vh>qHg=t#8^n5Skv=q`XBn~?xrs%*NYx9Y6_`}CbSPgr<=m!"
    "twI5g84;995p---"
    ")?|$%MoG?!dR^R0mL^@HbtDUY|`D1M$E2dX<rn{(2NHR8N9YG4rp@fXYJ+&``@+e)=dQo4$@Nn_kkF575ZV>9|&Du_bDJ$%(QIhp"
    "k%**b<RihgSroj9C|=07Qe<Qz4g#U7!NWTk*|_TH1b^-"
    "`~YIm#jAsEr(d@C>G=;qPhu2wfhMI3&mwWYQX7TZ*`K#rJ(%#IPybjWzXcqB4vq`=!AU4jhWC)hc}}w5vv{qw6c5V+N7af8;%i*X"
    "k&m&#BCLF_T9JQVm^zK?vYsTUWj+x@AV;ggP=VMSLvN_weEGss0QI>T|<d@2b1VJSD}#Tbjw|?0@Cr$QVA(Oa6IpU#stm`wnU`D4"
    "JdUh++$geOI_iIQUkwG24Z{Ckn5ML|I!lA!Y|}8mMn)EZ;5yX$zxT-"
    "190+x)#6+Rm4`v4hUNj4_)5eboCesbpt>+AaH*@UfUbU!!qkP4{r9P4Y%V#9kEcTrepJxsM=;BD?JLvh+Rbi7PLCe4Q30}JSDZS5"
    "%_Zv)Ti1WlIRv++E~nHuKZlXX*mNBQij4oRIJF6j<@9>#=RI3VwAPGa(2tf2pYoSfD5W66rY$~KS<g$c%hM+j4n#)BYy1reN)dUR"
    "n)szU6Yc4pV_*CB$%}n-"
    "DS}<IuQ#8?E+*l`veD`8j0e6~nAnGxBI0%XTl<wjd@53Qzg704``X>!4cC|_%)7~KDwa}^(Qx%Uz0k{Wyis-"
    "R*)dZ8^iqVp>FS31qe`=UN|>5#vTat3A!-anDdOL<H%(nBH^IxKqs*8~3Igwbv<VDli2d@CCHB&D<&FGWw}!bB_<h^zA6|-"
    "xFT3CEX|Gy&&wGB=rUNMj4OM%elazHHi_7?^uR9y8#|KIed~fj0I7&f*?La#((UKc)X=ZpX?1aA4#bqHK-"
    "7h`w|Lu@U!OW_?@#(YJ44%i^oVrgto?i*EPcB8QZRg!x)kJWjcP0%2zn0~(PcB8Q>yG`DT%7V+pQ9`doVtbdQ_cY>Mc^B*4d7Az$"
    "u!rss6PA|gi?gwcDnt8YFjy0h<`5;>u21)*zXKNfL?jX!cz)DY{%Vwp`7|J0r5{!0JPoC-OhNMec1_yejkZcBxT^Cy5v;Y`Y>Ns("
    "yLmmB1>ETi&{F2QZTS=Pd>a+`|LyAmrn@MyCcs-"
    "o(vF55&Eh%&8)_6h{NZ(xxB^S%iF$Bz^PD5LBxu^`0`R*^k+4{RjdZGTYOpw;rHfeot^EV&EMJKDz(=2FObNAvJ}Lvx=-"
    "TafAvA7a<tJqY-d3;cM_m5u>v%EC#0#CgC>dm&%a##tH|_xC%*qt{K|o@X?etfcUbRPcOm52p_Y@*dcLPlzAQ9Pmu=Qv)B`ueQ%-"
    "ssIp~qq9XVj=!Z3K?!FW+CqXAhd>8d3*;L<=!nha?^Nta3h8OG(WVNfbJ)QN14+pKbb%qNxI8M49CPQ^hv@}>AmGriJ){zxr<uhq"
    "umjlyy?e_d&ugI+C>yJfZhMR-H3q~0w{E79M}yQS)QxA5;@h)YX(xHrdq>XLZ8?-"
    "+<Jn(q|6<$UB<TuF?2OvE=Qm0K0(WYbJfJ?}2Hx?*pDTVHEA@0C^8d*y|CfTTW5o&ybm%&f1hb@;Nt&#m3*tzxaY6nmfC;yTNDht"
    "yo}kl10fl276lospwk!bR{^>|K%n+DdsdG?oC3{$|644h_PyO*RT-"
    "V{~*2EY{*L0k#YYuEsL3#cGrkO!}*e6eR*OO;Ah0WEG6biTM79`6)Get2>`bgN02q+){7~X9e_0`I9CZYPLj$XG3{CEqbjrU@Zle"
    "t6)_AEPnFBAXnY$bh()bn!&3HsHI@C1~x+X<=o`6qv_66qm=J~*SieM)0qys6y(-"
    "nM@43843TVY%cw<c06A7?>vAy=SUUY~1IU3oTfUBgy3)zq8!fQabC#(!CYg~eoiUVUV6f@ICOSt~4+dKQKDvLJ(21W{LPmUIb_Ew"
    "1pu3%Mzp{nK96S?xDJZUj9aLmf`l&(M#@phbs3RwW(}|9>6hs20fvj!bDAERH><4eI0zZm2lL{mWpt8DI0L?IS@OeQ&WEl-"
    "%fF_pZ_D#aF^(nQ%64byls|^{3T~8V9%7Dg@@vP<v|3&^|XX=iR?_l|?ghyQ@TiA5KrC=2(KV-3D#PCp-"
    "&ax3|fY`qBo=!B|+?CD$DCpBsmV!Z`VPS3-"
    "v=~?{sZ}l1WUx6W<?CH3ax^$PKR}XhH=`{DpKvmt_o~A|$WXOJJI}pV?gwa+btbf>;BysxT(Y>-"
    "gf*SI4rCLO71&a6Spy!KD>Gt~r1w#{{pleajXT3y3M%WcNmxFj@Hmlv<vQm!uKcS4S_%>yU;~6)9fub-"
    "9H^vr#DJ5$^7ywda>i3h&0TLX&1-g5Ff*R+hEaE6_i5s<(pT8oPUi1@=+m;}ysn9|F?%Dm-"
    "*FMhkbT)LU>KJHM1q3lGzzkQIpb1p9p9WowK4-"
    "8&Cik`3!IomGVoP<;HT|kM+O~mAYL$j?Ey_7@JH6c^h~vG@zwUD(o%lXo6R1u!HD^eZz>f0fwmcW&yV$et<;Strb)zG9nU_OsZj7"
    "o*=FJWUa4I!#5Mj^Kji-^yphN4p4af2G2oAF-"
    "9D$`TnYQ8R;j1@QnA53zxet48|HgjLn<2lv2NHW4E`jx!r?0TMyu^`{Vv<_Jzg;O@%RJUw7=(sfk3oTnk|9xLk)M#ik4x|ueww;="
    "wl5fAkg6}e~0%88$o^rgP@CMta}!0^z;E~2yXC1VNu=Dee?j;1Gd)-"
    "if4`7dk9Jukbkd*rT$2FQwF5~Fzo+@y7w^r*NlgeEhLI$8*=;%!qohEsZRP+o`{;yNBtD}fVKOTuaH!VQ_$ik3#h0+pUpVd72?s("
    "V*`ys!+;h)Q+0TzW_c&`+BlyJL$f~GF91K64RD9Dc(OLN`(~O$#?*l0bPf1PZNeJ}W|moSwT=_2$K0Y=pR4*PfiFyLL9hr}b*!!*"
    "t8eZlPGj7|Z}6NIP+iwSB!XK_3vTLfW?#pxgPXIoo*w0gVxED(_yq@9U}>FR)T=qFztXPZLVssJ30M81s=hM__(N+OsXeu-PB?6P"
    "E-"
    "u7B8Gj&ch4VvkITMa9^0AyrGUIAd)*G|PwA^QWEr9Sxyp}uSZ)IZ05bh0DaI9xg@CUkX*MwZnZFc0>KkK6!hOBaDJnZ_GgqV@w53"
    ";S49#s#-ig=|ve27xV#@Oq~MiT=3VQpH&8hP=bs+wBtSP%<ww-"
    "Tu#{YE(x5dMg_+$Xy6-pd!_2PP^drIpM(L9ep59V}%0UJAOx7=W6yQH3hE-"
    "soAu(!E6bp=yEI@4380usFVxd)q*jUh{3Ju+$WY?!2I?t?utK6lv!Lzua!jRt;9Q$K!CBs1fk-LP+WZCZzaT0&b5MMxKfG=yhz~l"
    "t)BP^ggPm8F^Q9Rni3{XfrnR_1YK9%U>+l(>C+fYV`M#JGhSLoC&gpmElNd&j3rM4mG=laYU`CwBlGN{S&vsl#9nTe`K_<k9+M=0"
    "fkZnibQ}{(gK|I%URcQ>j>v8mS%=J$|gi83#^vX41Lw9toK^L(-"
    "doL)EO60*qePefpFTYq<>I?@kAbM+jqV=_R;u536+byQ1gyH>np|M?$b^eNfJv(Nr0a!)X*NTKa}IolT&(F<uMj9tN|zimR7)63b"
    "!49{^jal^&RCSHu{P^V`w;v!6ZP?D%c78wb&w0)r5Z*u{boOgtt2F9d6X}6&sC)c(lnD`~<kHK~8u0to-"
    "8Skm+;HwmN8wb^*5C#x&i*0Ga@Abx4tN=5~3J3T0w@X`tL=bmeD>PAOJ-gr!L;29W?WkY0*WM*PsA4iKrt#hNH4&=Ft<!XMEko3#"
    "M<Ly9?;RmTM4r7DJhQajkQ@RT9k7-hvMzKPo$bGOBC128DU-"
    "p~nLf{N8ND)Rkv){DM1Lpc+WIjNcooMzw)%uAM>I)?=2C*Ik3wXa62d=|j}P0ux~vB6XB+WV|K@}1Sp0z{VyTki8f&0a8B4BQLE2"
    "p6xBHvpg;?bRHuPM2auxJn3(f-"
    "vO$aksF3rx*^pp|9l#{jvD(p6s^PqrcH%yr*wa7y&vOx*S+qO?aO6mBB&iZijXK8a2j0nQnmXZ6?v#=#@bCAtAcrlhv5-4hS6r-"
    "FP?cy;DCcgLawb{x<>hv-"
    "m6Bx}ZDkRvv{o*Soy}+;uI$=#tX1kn%>g3n(KE?$B_huk?4q6@H7m0*<||^qFYLa<IH);lSkG;oluS%Y=hWXqR2r_?KqQf>`697q"
    "iB{OxG@uiFY-"
    "}a&jcTxD@|AODrc`(aeflO>{7HDOkx%0skv`Mn5sJa4%^0i=`nBh%OTj2iO;C7Ild3cZnGdV!HjNcSjTh2)aRcIJ#a;2Y?IsXZqg"
    "$u^e0OJ?>^4bVFaa?_*^9&crYoj+@?o#oC@<wgw1v<K3{w`%2(C7K`bTUg8?;bh-oI;noq*4Y%$5`br!PPgDC;d1l7H?}r%$-"
    "Kg6Rs8Uz@CuMScphx|mx?gtK|Hq#Bjk*~H*{D~S+?^|C^H6ui73}eQ@H*{0^(c@tdBguMz{kRdb>W0T8C|Samvb;UolJZw?p`?>@"
    "r7PKyd!r-Ux$0rMNpRovOK)i61-iiJlChxj#K{ip}eyzH?P%fv7mBi$ZNzEr^?Nj;%Ck5UJ9UcEbC-WF8OBes4=^+jBZEJ%Pr4s("
    "pZL?pFTZ%^PjxruP!bf+D~KxHrX>a!ZJub9W>cjZlels$^J^*|L>KTq0vY`d>2=KgV0LffSoS4D`1{Q)5+!URPUE*!1{wRBEP8=%"
    "yij=X=qWUtJDq%QBA|s|B!;0E{#C;IkN)<&z>rCHT~o-"
    "w$4cU{8ACq<+2VuE<slJkJSi?KPIG<e00Kkcc;S!$VMlak<(?i0oEifMd)u;_uYXI%f~mmj~Um+vu9&8J=3Y6=`z{`Ke~jMsQ!iI"
    "xdqFd%5=bVsRWuEdphSDCJVnuqC#_^6Z%BtBL>ezod^Y~7U*QDZ7xAQr8!fJNFG(hxyBn)x(v2Sa;~6b*&CIt)ZVOuQR${Y!NOyi"
    "q$<r7_>EA)98^i^ux)^jsFSw}Z0njWK*eS5S^}kN6Da$Wq0L6YSqVhk%{u%#>DJ1CZq+PJ&qK4}rGEKJ&X*KBgbP(~_++B^Zn(1%"
    "<7-itG_$UQF6T|)SLUv`i1#|Z{H?jhm*4@i*JcbaT^d`!G-kI~;-"
    "atweVDB!u$|A9n3evm>Ww^2YP!e$VOQ;6m5`{i$jy)$GF>u(bU=sd4Res%SiZ%!!z+l|%TYX)si5i7S%sfQ;pWePZZJh@j2v!bP^"
    ";X^a(SH;?^;$u3zS>9*6Q|DPOd-"
    ";a{w%pWf0Uvtm$^|b@aQ}zl|m(YI*^_P2KA73iCrO;vNfjy4=>m%u@5arD5?<Ev2p6Eo^Z0rFQP99pzI_OsRqX;+9)8Lb}|+8Re$"
    "iK}1#%3*)Rb-Vh+hxwmj}U4Wl%sl7pKj>fxEJ9uPJ$UWj|(&q)(>*^2^A+M!{JoPuTuHx2#&FPn$9^uFuH-cE8D=xK&+pL4mvznP"
    "6urEAYkDWX8BO#0Eip%6P{>QiCk7vmp@J`V7KmRO&0DD|l?QzW>SgyvEdk(G32y$ehwYc|0@84b3U@Tw6-"
    ")kF=9#4b~J)4K6B0(PF)(Wt&_e}iCOGA-"
    "8eCu_w33_8_3!)0D>$?UtN{j!%OYz5jv0pgQPFww_opBpU(!&51`vX=3oa$qJJL8$|L+v#Vl~8opm!cqv-Ub$GKGr)qP+N|<pIhj"
    "jvj7J*F@V@f0a9_j72ljpS9onMRxkVm4sr}Yu~XCnOUha^M+-~J(n@dC9W!;WB88GLs?<CL#{Hd6Feo5br-Mj@ww@N+)Zfg$j#~#"
    "eXSOaq%2Bp%3bMd_9lfYm9nIBwKkV;J3eQc}F2q`W4gmhB+Ol?U)tuAM`i5qMjXK@HYee-A#~)5xvQBF2KhAag>N|KLU-"
    "g_#GNZsB-I`0%o<DvkT2T9*AN|+|<By|*b@m5gT6*})^(%cT-RQ7U;M>Pp|9Jc{ZP;_k7l~>$`?~8pZG4fEibof~ruF?kvvxaFcg"
    "VzGN4Ld#f#(-"
    "}3KH}|ZbC~IPgDAcpWtgFK22{UMn_W}IO~z}ylikOFn?Nbr9Ox~Qw^vxN#qoMECfEE^$h$vt88=vEOinSR{T^o!J5%*`y(vv_~bT"
    "HPz|4%X8}AV5;BJds76jb!EHYiN9p<5(HLc-"
    "p6sa_s;To*HPW3Yk2#f{&Nh;c6&)siet@jnPsBIhL}wiS$NwGvKs+DKb}GN>_t_9j6G|#({H(2jEtlE(S1gMmtwb%5wu%4kQRKrJ"
    "wKPiKoX_#Cd(h^6mmw51ozJPL@lyw!!^o@`p-"
    "N#;Dl9?ioi3ss6}MdYA!(q^%jO*{&9DZn_<3sh+SaF+y`v5N^m35y=gF1G0ox<p>y!(V<WN0s!O5i@<U0#vhuFi@T5^<3U5fAQN-"
    "v=JDK{8A?{&6W0g@>lgdW`1QbtFq`8Tt#<JQ5=DVa)-awM6Wf-"
    "F!nMK9`AdrexN+4iQhs<p&Tc*GQ2tReGVr#1k=A7Gn$YF~*zKO`nJ`3-"
    "37gYn1FCYw4^9y19E#X7gd@hzj!KOTQTYxec_wRqGt*S$`Ar{jH5oqaO?K-TR`=!ZE&-"
    "BY!fJJlA_FInYD;9cyK3I%_lZThJ*Rm(imTL`&)e&jW>`e);hX4A20oTrp7#s5A>1KhZ@={voF{41Uc>kbNOAy0>dKkP00P5e<LB"
    "mTrEtMT3CXYE!Vjo40jy77HG8&IH+?(&k$TFn1=kT{#MXDsl)_|APc{%9`SrRvh`Ryz89uP0ycvr**<>-"
    "rX|8!+IHZPjT%8~G%P@0{qn+J`~EfIY1a0e@(1N3It#6{w0_CZq7R4o_UJ3#)LSAx?#YKgzZP+0@};B2t3;dP}YF<5=5fe(zy33j"
    "EP^@SXXrQh_taxj9hX!E>HM=y`4$OaX#Fz;*i~aXwe-crQ*%yTT@OKe<|L^X}aH;puIy6Zl^;EWq#wKXH<qsG`PRf`}j0XoL%Kr@"
    "az8Q|;{Btvv2N<wf3y-"
    "9a!+4}YKk^*?8mkP0P@kW_$$@c`uK(^{(60iADS9WXMCm&t6BY{%(;6>h#4=z*J#75&3hhbhnVCFF&;`{i!$BDu%~y^tL&T{|gw@"
    "zVu1bZU0bnr4d0CkxFyF*98JY_-"
    "7WpInRfCH{|G=;JcBSUW&B_#}=OJ!1`^CWIF}@x_1UP)h6c<6%ga;A*<rSlFXwK#HHGW$>A8*OUg<#^B;-"
    "s|sxMcx0}Dt+Re*J?J7cdORXkC2OpJHOKuVWU9w9MzQ|5%0`H~i>5~|1{6O-"
    "tMH>{<xsBu#8lZwSXeTH#m`j@upKygrrPMRuLCSErx-"
    "^3%z&jT_#b!WEEQ4ge4DzPSnGs)zDMJoXiKrEe1#1#e(oAz{iNdkEat?in4D;3-"
    "Tr{}<cuXO&823{_*sMN9hl2u^W)_^V>X?&7ncz<0L0Hv3z8pEmQ?K%hT7<ZMXrD*fav*YE<y71Ub!@1Cd62spmIN0FgrRt{Ol})v"
    "6HJdyD*Z77H*qWu>4*g5eSOC16{R9a|QAob!a$8g%1+0gVZ4vYF5D>sQE3^p5D<CZR^&bW9gYU1IEu+4eUkrHquT;J?Y<ypZr~Pv"
    "eb4J8|)ch2k(UBvBkZ#?1=I62j|=<W=H9Ujl*saO}qRIAbwgPJsSB%{5+j_o|qlnjb^mtF=|Wz#?Mz1Y^v2yV*Ni>E6TCZaGhzht"
    "@=A1))tMfZ9L=GQ<3AR56HvO*7JpeDRRwWP)?3-"
    "RO7&*2*OpoJ{459hGR74DIF+&l0aP<o%>S}Su_%DWCgNov~*}=(UB2q<%V)(bnqfTwDluoS^za~Hq&$i1vRwXIiL&|bGgl6c@a^x"
    "rD?%6ddw+%q6l0iWJn=j<HaS7%n0#w)PSr=WwyO9vIs@1C;m3mW}_;bd*A%sFkCLQJWc#n`jR#)^_8ie{uY*rGg=&s`xd|mHBY%^"
    "cbmjRAzwS!U-9EXF$RgVk59yOIU}ZqU(0}h>uBcpMKpy~nIZ~!#X2I=2y@ldG-"
    "6lWy+;P6z&+2c_67);BkF3y5p|kLvtE+|V~uSG1^I9!uwV}GYL0Qo$BDM6jZT5Xj*pRt(rIigU(c?!R+uja&&U4f#;Q-"
    "o9Kw2z;hf_2EDmLcp5@1^u+f|`9GHV^vpzgg>7s;M_aPhjHZjsa9C09_GQ9I9EW5&u{J<o33g2HNbYiDDz?hj@h4)FYEj#F(5~!?"
    "CQ2F1Oq#;FsflqsZoSQnZU^@E%gi9z*;-myMVCSXy@0I@O-0iUKF_s4>h7mI-"
    "O|TmMzMQNpzVq4KQgye}uH5OW`K37O_NlNjv(|zvnWU#D7g;Z%kWQ;Q=70~C)3oN=G89wWyyt*pAD@V3H7%OvU(0}h>uBc80jGpD"
    "$^lDw1?GUsgt=<B#GekrOt!_c!N~VX+W-"
    "J_JZ%}^Q)P!FujrF%hmRYhJ>etdSVfpG_zZ3q8q86)rGek78oSJD%dSCve)@Ko`7K_gqQD$lCl_XAmZ$aj_WX|%h<!HZIM%I$1cb"
    "k!xF=4<cjD=^`(xIvJn=lj#S8*-SZ!HhHEmWp6Z7wf8_qu`=i@%^e%1CH+f*c&!`pNj(ZEv-r-ml4C%#LT`ln-"
    "#q}$M&KmT&|ulhRrh>gBt&sYJ&Zf2j3IFi-@H%h-1`3#+jKYpK3mu<5C(DMsF6$$3>F6S_&pN^I|&{h6-<T+p=ZuB~;0-"
    "_CNKVMA_#mNZ2Y+kr9EWMBU(z}S*#~@J(d@XU`U(OwX=8Y1h&2Siyy~+)<5K(D#i`R1Qgje>o<Cd5BQv9r$Da};vwK{9OnI1Ieug"
    "Jv^;-=S0iS=e3z8L=2XqzY)-"
    "@xRfV*rf!y##7bp!SUFG=3!SW#a~RJ$y6~^{S;q$4nlOQ$fDd917AX9{Zr234||oG(Yo9Kh{VO?Q!K1Lu8^I2aA{@D0>3S77|g9?"
    "gjc+6rdYHW=a)H*C&<SYYW&fkzewaS^l#&MkfE6!-"
    "|=q6?moNsiBT0WsXm<JVR!<n0Z<So>{Z2`jSr)%^7(7%FSpo(*&eN&0m~G!57ggb_+fiF_*(hO+cYDB8_3j%ok8b!nyK`f+=z(;u"
    "hS@1}E{;!i15ulu;md1)1W>1sD|&k@>ASYEzHJ7q88l`i%}7jUrvWf1RoUN+P~1Y4OeaW$OjqI@URrtC>NM>O}KU3skWZCc#11D("
    "}NYo^p-vlpZnPqSn}_Q;F(|{np-"
    "%hJA{O@FHz=lpH8_x5L7J*}E%;84HzQ+eEI1%3}AOcxvOD$Sv0QtQ|CbJeA1STx~j}e{ovwufqo8V(d@iK&XVdZdYl45^JHH9&##"
    "5F}g~-"
    "$6tQmjJ+$h%}A((+9r$rQM_t?7CYNB@xKu7p;l|C++l;Aj^||~6$zD4H?4Pdl#O~LHCx*KT)7kYzLyi9PbIWA&8!mo@h<d<%u6Ii"
    "h-"
    "<9vw+q3Hf=G<5YTD`TGhGVb>)HAZ;iU9U&@bjb)`3uo^s?)kJWX$W7GJ+Mj3U|`b_eV^3;+9U$ckB?Q&CZge;#G+EAf4q<rGyqIj"
    "HvfJEh{5QuzcgcRM^;yVV)@!Xx8oAqw1l&<iFHO1Qbwgc~5RtNiXhim+u3zP>+Z(co(=76n=0i9DE4gAt`w!`U>C7_<Knk0Xi#iV"
    "l+sORG@Tu4q!7{IO43U(SFJhX}Bgx)M;SkX3^kUgWyLMl!;P@J<}672&<wCtqYGQhpVvJPHctvOrVeuL1TNyAm1HW8HWs++6l|u="
    "tDVFsZNvWNNi5PD2wOfoQv}Zx@4J$PJpUq8Tq0(pq4%_@+}(j`Y0;AJVd}^&<2aH;YdPN`xw?zJzmi`MYW?;i=X=!V4}~X7QQPQl"
    "V@aSk!wmH__&AG9U?c??d(sO<$b_I29VJP(!ut^$H}zwW$_1pxPbs6-"
    "=>AL?>NYh3bHihVhd~dz8_J{E+p+og|(nWD9UAG}eGy!d#RgxEjMKN=SfOB(A6Dpj}zr(w?|>$k-"
    "6a*3tx(3Ryr25eHoxiYJS%01nXUJD1!OR}qb<E)>x4+gb0wupSq(3iMUSmW#LIi$m48c%;9aHbQW0&@G_j!i*#l?bWnsTYf+Dl5Q"
    "RBoIMw|sH?V%2x);W7o=IRG+j)S()^3?;rydmtG-L-(Kk9?6@2M9h{Sot?udD$Zm4}B;F`qUUU-Yk-"
    "A=dX+YQrzf=Gy+9L6Juu5ZOo{e!k{57@Bn-"
    "(EZcfk=F7_UV3~r07#St9=vr9_%;!Tq3c|GlueXPBL+j=%bD7Qh1l47(7<gP)Y+TBC&7S?egXQ^hWGK6Um27)~(zb54*m7^9>+~M"
    "A*`Aas5IhOkOGx(7t{GeHcb~Exen(?kLP4h(x$$oms+PH18rYxOikWLzrRL^I~C!KqJ1^mW$Ilg=iwOtdPiNxESD!>9W1Q`vc$Hi"
    "_D1T+p**Kc#-&QA|bV%-"
    "lm!9Nd1sVqCfC!uWtfCBeYO<xYmUNi5(Zk6$MIj=J>Ki#1zf9DwObSIy@r0nCqn2ePfJE>aJiCA!!+`Bq*(}-"
    "WetjZLz^OXuOLJNQqEY1rxfYCT_3@TBv1s0QF-"
    "oNlPrkF|ygs7F_R~V;&z%{~Xzbe+#UGE)s;lo))&}b`(KntAplMWq?YAt2*Q?gtu~W5sByP@E7&D2pdE8Y>a6zf&nlQ-WssIST1E"
    "`RLxjK>&<|Ux<WG->2co3t+KY4uRIuYAF9Pkj?55=Fa@U#a}_2nj3>fMnVO?G>{Q)~pR^RxHY-"
    "377W?!>Tx)4@HUC;V`dh~`r}ir)sFCDLpes=CMJ32pdywCVM@if9wnu?~&*Ir9BaYyT-"
    "Q4b@Y5<mx=%rkn^rYlKWct~N`F6QWg@ZWCHr>`AmB-"
    "7j?ti(}5xWkrLDfGUaUg51q5%Ef)M3Ra0g&H_b^H8x*@2wc13zMY&vVN4&qf@^hP}@JNiA0Te2y=DYbW#%M;yndJ&tE;-"
    "~V~K=*I&(6eEqhg95^8#)3D#x2Q(yJHkt4N^mItC^M=PO*A;-"
    ">df_^x47e1Don&AEkm}w{Y|A$kLIeg=SC+<ezd5Fa`FRYUXqrJK+I(oESES@DS$W8?o87GtOcX+RWcl-"
    "$t9WrBj#!a?75p?C}X=AF{tc2(}`Ygr?5P8_6tDF<0`O3%5ZeRrHjIS#sD<6kUPLsE99W}6pagV3S7iA)xnf%c@sofq-z8-"
    "Nl<PU=i<ck+sGwCD=$A_{R;0-"
    "!D4EsBSlOPSWVKBx?a=)i%MP_WUgmcp=$wSJ!uqW>W)>!Bel~Mx&!|G@as!W<}zfXPWNjzDxj{IiYF22dRnCI|Ag(@ZXNraRYwOB"
    "qVvie-"
    "U928<shYMADJlJ2>Dy?jV3(lRqn8^?}>@APbU@H6_?pu{EyP3^Mkf2ecCBtC|CUBNriXSUT*t&Iu<DYlI(slcpmt!xy3%5RAg)RL"
    "fi+@Tezp+wU1aBS=RPlii>?Zsn9yqIRB<zhL?tA;nBGLRp2|k1_Y#n>^$02USG7Yjm`DkWWA2(sgq_DqypWv_dj2VqfPvr_}I5oO"
    "8kRKg|uaNN4XL=%jfB&!T7y9?sdAu?!dP@iU9<rI9nYi#3gc*l6NBGA!|tIiL&yTzwf%0wO~8w`44#t9#Vl{&K3B`4gDZt0u%*)g"
    "FWjO8230FQgOE(^!C-Z>PX*n@Ntfw_jLM)lZxz$J+BcRvBZrec6ln_rv*Crf^~-l$bYl&P>Ma&M7Xze9V`ho!-nXvc{Ik<<r{-"
    "aiZP%fk@G$X1!?p*Oi)wdGLHmU#tG0oJHX9=<4v;PioB~F#hWadk1K2K$`+n-"
    "xfw7e_MqH(+>;w9xTBv1fpY9|OZ}lBjed?0(7WfuTs(G_0%4+A`Ghbl+{+^9VVY58#!ZUH6+d=aN@RLz^v!z6MmW8*=3$ef5l$zY"
    "UGu{tmL*#pgme9|7QgYh2-2;RLn&zw#cv&ngwbn#OS#22lUOvF2Bf42Yd|{3Sfo3Pbc`XZ>jCukxk-"
    "^}k?o76s6YWa{^Zk%E{r+yO2z;F-~WGXs}pM"
)

THAI_DIGITS = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")
PLAIN_PATTERN = re.compile(r"^[0-9]{13}$")
FORMATTED_PATTERN = re.compile(
    r"^[0-9]-[0-9]{4}-[0-9]{5}-[0-9]{2}-[0-9]$"
)
RCODE_PATTERN = re.compile(r"^[0-9]{4}$")

# ---------------------------------------------------------------------------
# Thai telephone numbering snapshot
# ---------------------------------------------------------------------------
# ตรวจทานข้อมูลอ้างอิงล่าสุด: 2026-08-08
# - โครงสร้าง/รหัสพื้นที่: NBTC + ITU (แหล่งทางการ)
# - ตาราง prefix -> operator ด้านล่างเป็นเพียง historical/community hint
#   เพราะประเทศไทยมี Mobile Number Portability (MNP) จึงใช้ prefix ยืนยัน
#   ผู้ให้บริการ "ปัจจุบัน" ไม่ได้
PHONE_REFERENCE_CHECKED = "2026-08-08"
PHONE_SOURCES: list[dict[str, str]] = [
    {
        "id": "nbtc_numbering_portal",
        "authority": "official",
        "title": "NBTC Telecommunication Numbering Management Bureau",
        "url": "https://numbering.nbtc.go.th/",
    },
    {
        "id": "itu_thailand_numbering_plan_2024",
        "authority": "official",
        "title": "ITU/NBTC Thailand numbering plan, communication 8 May 2024",
        "url": "https://www.itu.int/dms_pub/itu-t/oth/02/02/T02020000CD0004PDFE.pdf",
    },
    {
        "id": "nbtc_numbering_plan_2020",
        "authority": "official",
        "title": "NBTC numbering plan (2020)",
        "url": "https://numbering.nbtc.go.th/Announcement/Announce-manual/473.aspx",
    },
    {
        "id": "nbtc_mobile_allocation_2025",
        "authority": "official",
        "title": "NBTC cumulative mobile-number allocations, year 2568 (2025)",
        "url": "https://numbering.nbtc.go.th/FixedExistingNumber/558/643-(1).aspx",
    },
    {
        "id": "nbtc_mnp",
        "authority": "official",
        "title": "NBTC Mobile Number Portability (MNP / ย้ายค่ายเบอร์เดิม)",
        "url": "https://numbering.nbtc.go.th/Projects-and-activities/359.aspx",
    },
    {
        "id": "nbtc_short_3",
        "authority": "official",
        "title": "NBTC allocated 3-digit short numbers",
        "url": "https://numbering.nbtc.go.th/FixedExistingNumber/198/205.aspx",
    },
    {
        "id": "wikipedia_th_phone_prefix_snapshot",
        "authority": "secondary",
        "title": "Telephone numbers in Thailand - historical mobile prefix table",
        "url": "https://en.wikipedia.org/wiki/Telephone_numbers_in_Thailand",
    },
]

# รหัสพื้นที่โทรศัพท์ประจำที่ในรูปแบบภายในประเทศ (มี trunk prefix 0 แล้ว)
# active=True หมายถึงมีการระบุพื้นที่ใช้งานไว้ในแผน; active=False คือกลุ่มสำรอง
FIXED_AREA_CODES: dict[str, dict[str, Any]] = {
    "02": {
        "region_th": "กรุงเทพมหานครและปริมณฑล",
        "provinces_th": ["กรุงเทพมหานคร", "นนทบุรี", "ปทุมธานี", "สมุทรปราการ"],
        "note": "รวมบางพื้นที่พุทธมณฑล จ.นครปฐม ตามข้อมูลอ้างอิงสาธารณะ",
        "active": True,
    },
    "030": {"region_th": "ภาคกลาง/ตะวันออก/ตะวันตก", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "031": {"region_th": "ภาคกลาง/ตะวันออก/ตะวันตก", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "032": {"region_th": "ภาคกลาง/ตะวันตก", "provinces_th": ["ราชบุรี", "เพชรบุรี", "ประจวบคีรีขันธ์"], "note": None, "active": True},
    "033": {"region_th": "ภาคตะวันออก", "provinces_th": ["ฉะเชิงเทรา", "ระยอง", "ชลบุรี"], "note": None, "active": True},
    "034": {"region_th": "ภาคกลาง/ตะวันตก", "provinces_th": ["นครปฐม", "สมุทรสาคร", "กาญจนบุรี", "สมุทรสงคราม"], "note": None, "active": True},
    "035": {"region_th": "ภาคกลาง", "provinces_th": ["พระนครศรีอยุธยา", "สุพรรณบุรี", "อ่างทอง"], "note": None, "active": True},
    "036": {"region_th": "ภาคกลาง", "provinces_th": ["สระบุรี", "ลพบุรี", "สิงห์บุรี"], "note": None, "active": True},
    "037": {"region_th": "ภาคตะวันออก", "provinces_th": ["ปราจีนบุรี", "สระแก้ว", "นครนายก"], "note": None, "active": True},
    "038": {"region_th": "ภาคตะวันออก", "provinces_th": ["ฉะเชิงเทรา", "ระยอง", "ชลบุรี"], "note": None, "active": True},
    "039": {"region_th": "ภาคตะวันออก", "provinces_th": ["จันทบุรี", "ตราด"], "note": None, "active": True},
    "040": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "041": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "042": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": ["อุดรธานี", "บึงกาฬ", "หนองบัวลำภู", "หนองคาย", "นครพนม", "มุกดาหาร", "สกลนคร", "เลย"], "note": None, "active": True},
    "043": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": ["ขอนแก่น", "ร้อยเอ็ด", "มหาสารคาม", "กาฬสินธุ์"], "note": None, "active": True},
    "044": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": ["นครราชสีมา", "สุรินทร์", "บุรีรัมย์", "ชัยภูมิ"], "note": None, "active": True},
    "045": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": ["อุบลราชธานี", "อำนาจเจริญ", "ศรีสะเกษ", "ยโสธร"], "note": None, "active": True},
    "046": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "047": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "048": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "049": {"region_th": "ภาคตะวันออกเฉียงเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "050": {"region_th": "ภาคเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "051": {"region_th": "ภาคเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "052": {"region_th": "ภาคเหนือ", "provinces_th": ["เชียงใหม่", "ลำพูน", "แม่ฮ่องสอน", "เชียงราย"], "note": None, "active": True},
    "053": {"region_th": "ภาคเหนือ", "provinces_th": ["เชียงใหม่", "ลำพูน", "แม่ฮ่องสอน", "เชียงราย"], "note": None, "active": True},
    "054": {"region_th": "ภาคเหนือ", "provinces_th": ["ลำปาง", "พะเยา", "แพร่", "น่าน"], "note": None, "active": True},
    "055": {"region_th": "ภาคเหนือ", "provinces_th": ["พิษณุโลก", "อุตรดิตถ์", "ตาก", "สุโขทัย", "กำแพงเพชร"], "note": None, "active": True},
    "056": {"region_th": "ภาคเหนือ", "provinces_th": ["นครสวรรค์", "อุทัยธานี", "พิจิตร", "ชัยนาท", "เพชรบูรณ์"], "note": None, "active": True},
    "057": {"region_th": "ภาคเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "058": {"region_th": "ภาคเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "059": {"region_th": "ภาคเหนือ", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "070": {"region_th": "ภาคใต้", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "071": {"region_th": "ภาคใต้", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "072": {"region_th": "ภาคใต้", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "073": {"region_th": "ภาคใต้", "provinces_th": ["ยะลา", "ปัตตานี", "นราธิวาส"], "note": None, "active": True},
    "074": {"region_th": "ภาคใต้", "provinces_th": ["สงขลา", "พัทลุง", "สตูล"], "note": None, "active": True},
    "075": {"region_th": "ภาคใต้", "provinces_th": ["ตรัง", "นครศรีธรรมราช", "กระบี่"], "note": None, "active": True},
    "076": {"region_th": "ภาคใต้", "provinces_th": ["ภูเก็ต", "พังงา"], "note": None, "active": True},
    "077": {"region_th": "ภาคใต้", "provinces_th": ["สุราษฎร์ธานี", "ชุมพร", "ระนอง"], "note": None, "active": True},
    "078": {"region_th": "ภาคใต้", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
    "079": {"region_th": "ภาคใต้", "provinces_th": [], "note": "กลุ่มสำรอง", "active": False},
}

# ตารางนี้จงใจเรียกว่า "historical hint" ไม่ใช่ current carrier lookup.
# ใช้ longest-prefix match; ข้อมูลหลายช่วงเป็น secondary source และอาจเก่า/ไม่ครบ.
MOBILE_PREFIX_OPERATOR_HINTS: dict[str, dict[str, str]] = {
    "060": {"operator": "TrueMove", "service": "VoIP", "confidence": "secondary"},
    "061": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "063": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "064": {"operator": "True", "service": "mobile", "confidence": "secondary"},
    "065": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "066": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "068": {"operator": "TOT/NT", "service": "VoIP (historical)", "confidence": "secondary"},
    "0800": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0801": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0802": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0803": {"operator": "TrueMove", "service": "mobile", "confidence": "secondary"},
    "0804": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "0805": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "0806": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0810": {"operator": "AIS", "service": "mobile", "confidence": "secondary-low"},
    "0811": {"operator": "AIS", "service": "mobile", "confidence": "secondary-low"},
    "0812": {"operator": "AIS", "service": "mobile", "confidence": "secondary-low"},
    "0813": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "0814": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "0815": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "0816": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "0817": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0818": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0819": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "082": {"operator": "AIS", "service": "mobile", "confidence": "secondary-low"},
    "083": {"operator": "TrueMove", "service": "mobile", "confidence": "secondary"},
    "084": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0854": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "085": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "086": {"operator": "TrueMove", "service": "mobile", "confidence": "secondary"},
    "087": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "088": {"operator": "TrueMove", "service": "mobile", "confidence": "secondary"},
    "089": {"operator": "TrueMove H", "service": "mobile (historical table)", "confidence": "secondary-low"},
    "0901": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0909": {"operator": "TrueMove H / AIS", "service": "mobile", "confidence": "secondary"},
    "090": {"operator": "dtac", "service": "mobile", "confidence": "secondary"},
    "091": {"operator": "TrueMove H", "service": "mobile", "confidence": "secondary"},
    "092": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0931": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "093": {"operator": "AIS / TrueMove H", "service": "mobile", "confidence": "secondary"},
    "0940": {"operator": "TrueMove H", "service": "mobile", "confidence": "secondary"},
    "0941": {"operator": "TrueMove H", "service": "mobile", "confidence": "secondary"},
    "09422": {"operator": "TrueMove H", "service": "mobile", "confidence": "secondary"},
    "09423": {"operator": "TrueMove H", "service": "mobile", "confidence": "secondary"},
    "094": {"operator": "dtac / TrueMove H", "service": "mobile", "confidence": "secondary"},
    "097": {"operator": "AIS / TrueMove H", "service": "mobile", "confidence": "secondary"},
    "098": {"operator": "AIS", "service": "mobile", "confidence": "secondary"},
    "0995": {"operator": "TrueMove H", "service": "mobile", "confidence": "secondary"},
    "099": {"operator": "AIS / TrueMove H / dtac", "service": "mobile", "confidence": "secondary"},
}

PREFIX_HISTORY: dict[str, dict[str, str]] = {
    "09": {
        "event": "กลุ่ม 09x ถูกเพิ่มเพื่อขยายเลขหมายมือถือ; มีข้อมูลสาธารณะว่า 090 เริ่มพร้อมจัดสรรตั้งแต่ 28 เม.ย. 2011",
        "source_quality": "secondary",
        "important": "ไม่ใช่ปีที่ซิมหรือเลขรายนี้ถูกออกให้ผู้ใช้",
    },
    "06": {
        "event": "กสทช. เห็นชอบเพิ่มเลขหมายสำรองในหมวด 06 อีก 50 ล้านเลขหมายในปี 2557 (2014)",
        "source_quality": "official",
        "important": "เป็นประวัติของหมวดเลข ไม่ใช่ปีออกของเลขรายนี้",
    },
    "061": {
        "event": "061 ปรากฏในชุดเลขหมายสวยที่ กสทช. เตรียมประมูลเมื่อปี 2559 (2016)",
        "source_quality": "official",
        "important": "เป็นหลักฐานว่า prefix มีใช้งาน/จัดการในช่วงนั้น ไม่ใช่ปีออกของเลขรายนี้",
    },
    "062": {
        "event": "062 ปรากฏในชุดเลขหมายสวยที่ กสทช. เตรียมประมูลเมื่อปี 2559 (2016)",
        "source_quality": "official",
        "important": "เป็นหลักฐานว่า prefix มีใช้งาน/จัดการในช่วงนั้น ไม่ใช่ปีออกของเลขรายนี้",
    },
    "063": {
        "event": "063 ปรากฏในชุดเลขหมายสวยที่ กสทช. เตรียมประมูลเมื่อปี 2559 (2016)",
        "source_quality": "official",
        "important": "เป็นหลักฐานว่า prefix มีใช้งาน/จัดการในช่วงนั้น ไม่ใช่ปีออกของเลขรายนี้",
    },
    "064": {
        "event": "064 ปรากฏในชุดเลขหมายสวยที่ กสทช. เตรียมประมูลเมื่อปี 2559 (2016)",
        "source_quality": "official",
        "important": "เป็นหลักฐานว่า prefix มีใช้งาน/จัดการในช่วงนั้น ไม่ใช่ปีออกของเลขรายนี้",
    },
    "065": {
        "event": "065 ปรากฏในชุดเลขหมายสวยที่ กสทช. เตรียมประมูลเมื่อปี 2559 (2016)",
        "source_quality": "official",
        "important": "เป็นหลักฐานว่า prefix มีใช้งาน/จัดการในช่วงนั้น ไม่ใช่ปีออกของเลขรายนี้",
    },
    "096": {
        "event": "096 ปรากฏในชุดเลขหมายสวยที่ กสทช. เตรียมประมูลเมื่อปี 2559 (2016)",
        "source_quality": "official",
        "important": "เป็นหลักฐานว่า prefix มีใช้งาน/จัดการในช่วงนั้น ไม่ใช่ปีออกของเลขรายนี้",
    },
}

SHORT_NUMBER_HINTS: dict[str, dict[str, str]] = {
    "191": {"organization": "สำนักงานตำรวจแห่งชาติ", "purpose": "บริการโทรศัพท์ฉุกเฉินแห่งชาติ", "source_quality": "official"},
    "192": {"organization": "กระทรวงดิจิทัลเพื่อเศรษฐกิจและสังคม", "purpose": "เลขหมายสั้น 3 หลัก", "source_quality": "official"},
    "194": {"organization": "กรมสื่อสารอิเล็กทรอนิกส์ทหารอากาศ กองทัพอากาศ", "purpose": "เลขหมายสั้น 3 หลัก", "source_quality": "official"},
    "199": {"organization": "กองบังคับการตำรวจดับเพลิง สังกัดกรุงเทพมหานคร", "purpose": "ดับเพลิง", "source_quality": "official"},
}
XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


# ตารางจังหวัดในตัวโปรแกรมใช้เป็น fallback เมื่อยังโหลด RCODE ไม่ได้
PROVINCES_TH: dict[str, str] = {
    "10": "กรุงเทพมหานคร",
    "11": "สมุทรปราการ",
    "12": "นนทบุรี",
    "13": "ปทุมธานี",
    "14": "พระนครศรีอยุธยา",
    "15": "อ่างทอง",
    "16": "ลพบุรี",
    "17": "สิงห์บุรี",
    "18": "ชัยนาท",
    "19": "สระบุรี",
    "20": "ชลบุรี",
    "21": "ระยอง",
    "22": "จันทบุรี",
    "23": "ตราด",
    "24": "ฉะเชิงเทรา",
    "25": "ปราจีนบุรี",
    "26": "นครนายก",
    "27": "สระแก้ว",
    "30": "นครราชสีมา",
    "31": "บุรีรัมย์",
    "32": "สุรินทร์",
    "33": "ศรีสะเกษ",
    "34": "อุบลราชธานี",
    "35": "ยโสธร",
    "36": "ชัยภูมิ",
    "37": "อำนาจเจริญ",
    "38": "บึงกาฬ",
    "39": "หนองบัวลำภู",
    "40": "ขอนแก่น",
    "41": "อุดรธานี",
    "42": "เลย",
    "43": "หนองคาย",
    "44": "มหาสารคาม",
    "45": "ร้อยเอ็ด",
    "46": "กาฬสินธุ์",
    "47": "สกลนคร",
    "48": "นครพนม",
    "49": "มุกดาหาร",
    "50": "เชียงใหม่",
    "51": "ลำพูน",
    "52": "ลำปาง",
    "53": "อุตรดิตถ์",
    "54": "แพร่",
    "55": "น่าน",
    "56": "พะเยา",
    "57": "เชียงราย",
    "58": "แม่ฮ่องสอน",
    "60": "นครสวรรค์",
    "61": "อุทัยธานี",
    "62": "กำแพงเพชร",
    "63": "ตาก",
    "64": "สุโขทัย",
    "65": "พิษณุโลก",
    "66": "พิจิตร",
    "67": "เพชรบูรณ์",
    "70": "ราชบุรี",
    "71": "กาญจนบุรี",
    "72": "สุพรรณบุรี",
    "73": "นครปฐม",
    "74": "สมุทรสาคร",
    "75": "สมุทรสงคราม",
    "76": "เพชรบุรี",
    "77": "ประจวบคีรีขันธ์",
    "80": "นครศรีธรรมราช",
    "81": "กระบี่",
    "82": "พังงา",
    "83": "ภูเก็ต",
    "84": "สุราษฎร์ธานี",
    "85": "ระนอง",
    "86": "ชุมพร",
    "90": "สงขลา",
    "91": "สตูล",
    "92": "ตรัง",
    "93": "พัทลุง",
    "94": "ปัตตานี",
    "95": "ยะลา",
    "96": "นราธิวาส",
}


# หลักแรกสะท้อนประเภทตอนจัดสรรเลข ไม่ควรใช้ยืนยันสถานะปัจจุบัน
PERSON_TYPES: dict[int, dict[str, str]] = {
    0: {
        "title": "บุคคลไม่มีสถานะทางทะเบียน",
        "detail": (
            "ได้รับการสำรวจหรือจัดทำทะเบียนประวัติในกลุ่มเลข 0 "
            "เลขหลักแรกไม่ยืนยันสัญชาติหรือสถานะปัจจุบัน"
        ),
    },
    1: {
        "title": "ผู้เกิดและได้สัญชาติไทย แจ้งเกิดภายในกำหนด",
        "detail": (
            "ได้รับเลขประเภท 1 จากการแจ้งเกิดภายในกำหนด "
            "หลังเริ่มใช้ระบบเลขประจำตัว 13 หลัก"
        ),
    },
    2: {
        "title": "ผู้เกิดและได้สัญชาติไทย แจ้งเกิดเกินกำหนด",
        "detail": "ได้รับเลขประเภท 2 จากการแจ้งเกิดหลังพ้นกำหนด",
    },
    3: {
        "title": "มีชื่อในทะเบียนบ้านช่วงเริ่มระบบเลข 13 หลัก",
        "detail": (
            "มีชื่ออยู่ในทะเบียนบ้านเดิมแล้วในช่วงที่รัฐเริ่มจัดสรรเลข "
            "13 หลัก คนไทยที่เกิดก่อนระบบใหม่จำนวนมากอยู่ในประเภทนี้"
        ),
    },
    4: {
        "title": "บุคคลในทะเบียนเดิมที่ได้รับเลขภายหลัง",
        "detail": (
            "มีรายการทะเบียนเดิมแต่ยังไม่มีเลขในระบบใหม่ "
            "และได้รับเลขเมื่อดำเนินการทางทะเบียนภายหลัง"
        ),
    },
    5: {
        "title": "คนไทยที่เพิ่มชื่อในทะเบียนบ้านภายหลัง",
        "detail": (
            "ได้รับการเพิ่มชื่อในทะเบียนบ้านภายหลัง "
            "เช่น กรณีตกสำรวจหรือกรณีทางทะเบียนอื่น"
        ),
    },
    6: {
        "title": "บุคคลต่างด้าวหรือผู้มีสถานะการอยู่บางประเภท",
        "detail": (
            "ใช้กับผู้เข้าเมืองหรือผู้มีสถานะการอยู่บางกลุ่ม "
            "เลขหลักแรกไม่ยืนยันสถานะปัจจุบัน"
        ),
    },
    7: {
        "title": "บุตรของบุคคลประเภท 6 ที่เกิดในประเทศไทย",
        "detail": (
            "สะท้อนกลุ่มทะเบียนตอนจัดสรรเลข "
            "ไม่ใช้ยืนยันสัญชาติหรือสถานะปัจจุบัน"
        ),
    },
    8: {
        "title": "บุคคลที่ได้รับสถานะหรือสัญชาติภายหลังบางกรณี",
        "detail": (
            "เช่น ผู้มีถิ่นที่อยู่โดยชอบ ผู้แปลงสัญชาติ "
            "หรือผู้ได้รับสัญชาติไทยภายหลัง"
        ),
    },
}


@dataclass
class UpdateResult:
    attempted: bool
    status: str
    cache_path: str
    message: str
    used_cache: bool
    downloaded: bool
    record_count: int = 0
    data_as_of_th: str | None = None
    checked_at: str | None = None
    error: str | None = None


@dataclass
class ChecksumResult:
    actual: int
    expected: int
    valid: bool
    total: int
    remainder: int
    products: list[dict[str, int]]


@dataclass
class DecodeResult:
    raw_input: str
    normalized_id: str | None = None
    formatted_id: str | None = None
    displayed_id: str | None = None
    input_format: str | None = None
    format_valid: bool = False
    checksum_valid: bool = False
    person_type_known: bool = False
    registry_known: bool = False
    registry_active: bool | None = None
    structurally_valid: bool = False
    reference_match: bool = False
    existence_verified: bool = False
    person_type: dict[str, Any] | None = None
    registry: dict[str, Any] | None = None
    serial: dict[str, Any] | None = None
    checksum: dict[str, Any] | None = None
    database: dict[str, Any] | None = None
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class PhoneDecodeResult:
    raw_input: str
    input_kind: str = "thai_phone"
    normalized_national: str | None = None
    formatted_national: str | None = None
    displayed_number: str | None = None
    e164: str | None = None
    input_format: str | None = None
    format_valid: bool = False
    structurally_valid: bool = False
    number_type: str | None = None
    numbering_category: str | None = None
    country_code: str = "+66"
    trunk_prefix: str | None = None
    national_significant_number: str | None = None
    prefix: str | None = None
    area_code: str | None = None
    area: dict[str, Any] | None = None
    subscriber_number: str | None = None
    historical_operator_hint: dict[str, Any] | None = None
    current_operator_verified: bool = False
    assignment_verified: bool = False
    issue_year: int | None = None
    prefix_history: dict[str, Any] | None = None
    special_service: dict[str, Any] | None = None
    source_snapshot_date: str = PHONE_REFERENCE_CHECKED
    source_notes: list[dict[str, str]] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class RCodeError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------


def configure_console() -> None:
    """ตั้ง UTF-8 ให้ stdout/stderr บน Windows รุ่นใหม่เมื่อทำได้"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except OSError:
                pass


def now_iso() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def normalize_search_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).casefold().strip()
    return re.sub(r"\s+", " ", value)


def atomic_write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = json.dumps(
        data,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=False,
    )

    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        temporary.write_text(encoded, encoding="utf-8")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def load_database(path: Path) -> dict[str, Any] | None:
    if not path.exists():
        return None

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RCodeError(f"อ่าน {path.name} ไม่สำเร็จ: {error}") from error

    if not isinstance(data, dict):
        raise RCodeError("ฐาน RCODE ต้องเป็น JSON object")

    if not isinstance(data.get("registries"), dict):
        raise RCodeError("ฐาน RCODE ไม่มี registries ที่ถูกต้อง")

    if not isinstance(data.get("provinces"), dict):
        data["provinces"] = {
            code: {"name_th": name, "name_en": None}
            for code, name in PROVINCES_TH.items()
        }

    return data


def create_fallback_database(error: str | None = None) -> dict[str, Any]:
    """สร้างฐานจาก snapshot ที่ฝังในไฟล์; ถ้า snapshot เสียจึงใช้ตารางจังหวัดล้วน"""
    checked_at = now_iso()

    try:
        compressed = base64.b85decode(EMBEDDED_RCODE_ZLIB_B85.encode("ascii"))
        decoded = zlib.decompress(compressed).decode("utf-8")
        database = json.loads(decoded)
        if not isinstance(database.get("registries"), dict):
            raise ValueError("embedded registries ไม่ถูกต้อง")

        metadata = database.setdefault("metadata", {})
        metadata["schema_version"] = CACHE_SCHEMA_VERSION
        metadata["source_page_url"] = SOURCE_PAGE_URL
        metadata["source_url"] = SOURCE_XLSX_URL
        metadata["last_checked_at"] = checked_at
        metadata["cache_created_at"] = checked_at
        metadata["fallback"] = True
        metadata["embedded_snapshot"] = True
        metadata["last_error"] = error
        metadata["record_count"] = len(database["registries"])
        metadata.setdefault("notes", []).append(
            "สร้างจาก snapshot ในไฟล์ Python เพราะดาวน์โหลดต้นทางไม่ได้"
        )
        return database
    except (ValueError, UnicodeError, json.JSONDecodeError, zlib.error) as snapshot_error:
        combined_error = f"{error or ''}; embedded snapshot: {snapshot_error}".strip("; " )
        return {
            "metadata": {
                "schema_version": CACHE_SCHEMA_VERSION,
                "source": "สำนักบริหารการทะเบียน กรมการปกครอง",
                "source_page_url": SOURCE_PAGE_URL,
                "source_url": SOURCE_XLSX_URL,
                "source_file": "rcode.xlsx",
                "generated_at": checked_at,
                "last_checked_at": checked_at,
                "data_as_of_th": None,
                "record_count": 0,
                "fallback": True,
                "embedded_snapshot": False,
                "last_error": combined_error,
                "notes": [
                    "snapshot ใช้งานไม่ได้ จึงเหลือเพียง lookup จังหวัด",
                    "โปรแกรมจะลองอัปเดตใหม่ในการรันครั้งถัดไป",
                ],
            },
            "provinces": {
                code: {"name_th": name, "name_en": None}
                for code, name in PROVINCES_TH.items()
            },
            "registries": {},
        }


def database_summary(database: dict[str, Any] | None, path: Path) -> dict[str, Any]:
    if not database:
        return {
            "loaded": False,
            "cache_path": str(path),
            "record_count": 0,
            "data_as_of_th": None,
            "source_url": SOURCE_XLSX_URL,
        }

    metadata = database.get("metadata", {})
    registries = database.get("registries", {})
    return {
        "loaded": True,
        "cache_path": str(path),
        "record_count": metadata.get("record_count", len(registries)),
        "data_as_of_th": metadata.get("data_as_of_th"),
        "generated_at": metadata.get("generated_at"),
        "last_checked_at": metadata.get("last_checked_at"),
        "fallback": bool(metadata.get("fallback")),
        "source": metadata.get("source"),
        "source_url": metadata.get("source_url", SOURCE_XLSX_URL),
        "source_page_url": metadata.get("source_page_url", SOURCE_PAGE_URL),
        "sha256": metadata.get("sha256"),
    }


# ---------------------------------------------------------------------------
# XLSX reader: standard library only
# ---------------------------------------------------------------------------


def _xlsx_column_name(cell_ref: str) -> str:
    match = re.match(r"[A-Z]+", cell_ref.upper())
    return match.group(0) if match else ""


def _safe_zip_members(archive: zipfile.ZipFile) -> None:
    total = 0
    for info in archive.infolist():
        total += info.file_size
        if total > MAX_XLSX_UNCOMPRESSED_BYTES:
            raise RCodeError("ไฟล์ XLSX มีขนาดคลายบีบอัดใหญ่ผิดปกติ")
        if ".." in Path(info.filename).parts:
            raise RCodeError("พบ path ที่ไม่ปลอดภัยในไฟล์ XLSX")


def _first_worksheet_path(archive: zipfile.ZipFile) -> str:
    fallback = "xl/worksheets/sheet1.xml"
    try:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.find(f"{{{XLSX_MAIN_NS}}}sheets")
        if sheets is None or len(sheets) == 0:
            return fallback

        relation_id = sheets[0].attrib.get(f"{{{XLSX_REL_NS}}}id")
        if not relation_id:
            return fallback

        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        for relation in relations.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
            if relation.attrib.get("Id") != relation_id:
                continue
            target = relation.attrib.get("Target", "")
            if target.startswith("/"):
                return target.lstrip("/")
            return str(Path("xl") / target).replace("\\", "/")
    except (KeyError, ET.ParseError, IndexError):
        return fallback

    return fallback


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    except ET.ParseError as error:
        raise RCodeError(f"อ่าน sharedStrings.xml ไม่สำเร็จ: {error}") from error

    strings: list[str] = []
    for item in root.findall(f"{{{XLSX_MAIN_NS}}}si"):
        text = "".join(
            node.text or ""
            for node in item.iter(f"{{{XLSX_MAIN_NS}}}t")
        )
        strings.append(text)
    return strings


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> str | None:
    cell_type = cell.attrib.get("t")

    if cell_type == "inlineStr":
        inline = cell.find(f"{{{XLSX_MAIN_NS}}}is")
        if inline is None:
            return None
        return "".join(
            node.text or ""
            for node in inline.iter(f"{{{XLSX_MAIN_NS}}}t")
        )

    value_node = cell.find(f"{{{XLSX_MAIN_NS}}}v")
    if value_node is None or value_node.text is None:
        return None

    raw = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError) as error:
            raise RCodeError("ดัชนี shared string ใน XLSX ไม่ถูกต้อง") from error

    return raw


def read_xlsx_rows(content: bytes) -> list[list[str | None]]:
    if not content.startswith(b"PK"):
        raise RCodeError("ข้อมูลที่ดาวน์โหลดไม่ใช่ไฟล์ XLSX/ZIP")

    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            _safe_zip_members(archive)
            shared_strings = _read_shared_strings(archive)
            sheet_path = _first_worksheet_path(archive)
            sheet_root = ET.fromstring(archive.read(sheet_path))

            rows: list[list[str | None]] = []
            row_path = (
                f".//{{{XLSX_MAIN_NS}}}sheetData/"
                f"{{{XLSX_MAIN_NS}}}row"
            )
            for row_node in sheet_root.findall(row_path):
                values: dict[str, str | None] = {}
                for cell in row_node.findall(f"{{{XLSX_MAIN_NS}}}c"):
                    column = _xlsx_column_name(cell.attrib.get("r", ""))
                    if column in {"A", "B", "C", "D"}:
                        values[column] = _cell_value(cell, shared_strings)

                rows.append(
                    [
                        values.get("A"),
                        values.get("B"),
                        values.get("C"),
                        values.get("D"),
                    ]
                )
    except zipfile.BadZipFile as error:
        raise RCodeError("ไฟล์ที่ดาวน์โหลดไม่ใช่ XLSX ที่สมบูรณ์") from error
    except KeyError as error:
        raise RCodeError(f"ไฟล์ XLSX ขาดส่วนประกอบ: {error}") from error
    except ET.ParseError as error:
        raise RCodeError(f"โครงสร้าง XML ใน XLSX ไม่ถูกต้อง: {error}") from error

    return rows


def classify_office(name_th: str) -> tuple[str, str, str]:
    """คืนชนิด machine-readable, ป้ายภาษาไทย และชื่อพื้นที่แบบตัด prefix"""
    patterns = (
        ("local_city", "เทศบาลนคร", "ท้องถิ่นเทศบาลนคร"),
        ("local_town", "เทศบาลเมือง", "ท้องถิ่นเทศบาลเมือง"),
        ("local_subdistrict", "เทศบาลตำบล", "ท้องถิ่นเทศบาลตำบล"),
        ("bangkok_district", "สำนักงานเขต", "ท้องถิ่นเขต"),
        ("pattaya", "เมืองพัทยา", "ท้องถิ่นเมืองพัทยา"),
        ("minor_district", "กิ่งอำเภอ", "กิ่งอำเภอ"),
        ("district", "อำเภอ", "อำเภอ"),
        ("province", "จังหวัด", "จังหวัด"),
        ("branch", "สำนักทะเบียนสาขา", "สาขา"),
    )

    clean = name_th.strip().lstrip("*").strip()
    for kind, label, prefix in patterns:
        if clean.startswith(prefix):
            locality = clean[len(prefix):].strip()
            return kind, label, locality or clean

    if clean == "สำนักทะเบียนกลาง":
        return "central", "สำนักทะเบียนกลาง", clean

    return "other", "สำนักทะเบียนอื่น", clean


def parse_discontinued(raw: str | None, name_th: str) -> dict[str, Any]:
    value = (raw or "0").strip()
    marked_cancelled = name_th.strip().startswith("*")

    if value in {"", "0", "0.0"} and not marked_cancelled:
        return {
            "active": True,
            "discontinued_be": None,
            "discontinued_iso": None,
        }

    digits = re.sub(r"\D", "", value)
    if len(digits) == 8:
        year_be = int(digits[:4])
        month = int(digits[4:6])
        day = int(digits[6:8])
        try:
            iso = date(year_be - 543, month, day).isoformat()
        except ValueError:
            iso = None
        return {
            "active": False,
            "discontinued_be": f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}",
            "discontinued_iso": iso,
        }

    return {
        "active": False,
        "discontinued_be": value if value not in {"", "0", "0.0"} else None,
        "discontinued_iso": None,
    }


def build_database(
    xlsx_content: bytes,
    *,
    http_etag: str | None = None,
    http_last_modified: str | None = None,
) -> dict[str, Any]:
    rows = read_xlsx_rows(xlsx_content)

    data_date: str | None = None
    for row in rows[:20]:
        first = str(row[0] or "")
        match = re.search(r"ข้อมูล\s*ณ\s*วันที่\s*(.+)", first)
        if match:
            data_date = match.group(1).strip()
            break

    raw_records: list[tuple[str, str, str, str | None]] = []
    for code_raw, name_th_raw, name_en_raw, discontinued_raw in rows:
        code_text = str(code_raw or "").strip()
        digits = re.sub(r"\D", "", code_text)
        if not digits:
            continue

        try:
            code = f"{int(digits):04d}"
        except ValueError:
            continue

        name_th = str(name_th_raw or "").strip()
        if len(code) != 4 or not name_th:
            continue

        name_en = str(name_en_raw or "").strip()
        raw_records.append((code, name_th, name_en, discontinued_raw))

    if len(raw_records) < 100:
        raise RCodeError(
            f"พบข้อมูล RCODE เพียง {len(raw_records)} แถว ซึ่งน้อยผิดปกติ"
        )

    provinces: dict[str, dict[str, str | None]] = {
        code: {"name_th": name, "name_en": None}
        for code, name in PROVINCES_TH.items()
    }

    # รหัสจังหวัดในไฟล์มักลงท้าย 00 เช่น 3000 = จังหวัดนครราชสีมา
    for code, name_th, name_en, _ in raw_records:
        if not code.endswith("00"):
            continue

        province_code = code[:2]
        clean_name = name_th.lstrip("*").strip()
        if code == "1000":
            province_name_th = "กรุงเทพมหานคร"
        elif clean_name.startswith("จังหวัด"):
            province_name_th = clean_name.removeprefix("จังหวัด").strip()
        else:
            province_name_th = clean_name

        provinces[province_code] = {
            "name_th": province_name_th,
            "name_en": name_en or None,
        }

    registries: dict[str, dict[str, Any]] = {}
    for code, name_th_raw, name_en, discontinued_raw in raw_records:
        name_th = name_th_raw.lstrip("*").strip()
        province_code = code[:2]
        office_type, office_type_th, locality = classify_office(name_th_raw)
        status = parse_discontinued(discontinued_raw, name_th_raw)
        province = provinces.get(province_code, {})

        registries[code] = {
            "name_th": name_th,
            "name_en": name_en or None,
            "office_type": office_type,
            "office_type_th": office_type_th,
            "locality_name_th": locality,
            "province_code": province_code,
            "province_name_th": province.get("name_th"),
            "province_name_en": province.get("name_en"),
            **status,
        }

    checked_at = now_iso()
    return {
        "metadata": {
            "schema_version": CACHE_SCHEMA_VERSION,
            "source": "สำนักบริหารการทะเบียน กรมการปกครอง",
            "source_page_url": SOURCE_PAGE_URL,
            "source_url": SOURCE_XLSX_URL,
            "source_file": "rcode.xlsx",
            "data_as_of_th": data_date,
            "generated_at": checked_at,
            "last_checked_at": checked_at,
            "record_count": len(registries),
            "fallback": False,
            "sha256": hashlib.sha256(xlsx_content).hexdigest(),
            "http_etag": http_etag,
            "http_last_modified": http_last_modified,
            "last_error": None,
            "notes": [
                "active=false หมายถึงชุดข้อมูลระบุว่ารหัสถูกจำหน่ายหรือยกเลิก",
                "รหัสสำนักทะเบียนในเลขไม่ใช่ที่อยู่ปัจจุบันหรือสถานที่เกิดโดยตรง",
            ],
        },
        "provinces": dict(sorted(provinces.items())),
        "registries": dict(sorted(registries.items())),
    }


# ---------------------------------------------------------------------------
# Automatic updater
# ---------------------------------------------------------------------------


def _validate_xlsx_content(content: bytes) -> None:
    """ตรวจว่า response เป็นไฟล์ XLSX จริง ไม่ใช่หน้า error HTML"""
    if not content:
        raise RCodeError("ไฟล์ที่ดาวน์โหลดมาว่างเปล่า")
    if len(content) > MAX_DOWNLOAD_BYTES:
        raise RCodeError("ไฟล์ RCODE ใหญ่เกินขนาดที่กำหนด")
    if not content.startswith(b"PK") or not zipfile.is_zipfile(io.BytesIO(content)):
        preview = content[:120].decode("utf-8", errors="replace").replace("\n", " ")
        raise RCodeError(f"ต้นทางไม่ได้ส่งไฟล์ XLSX ที่ถูกต้อง: {preview!r}")


def _download_with_powershell(url: str, timeout: float) -> bytes:
    """Fallback สำหรับ Windows โดยใช้ TLS/certificate store ของระบบ"""
    shell = None
    for candidate in ("pwsh.exe", "powershell.exe", "pwsh", "powershell"):
        try:
            probe = subprocess.run(
                [candidate, "-NoProfile", "-NonInteractive", "-Command", "$PSVersionTable.PSVersion.ToString()"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=5,
                check=False,
            )
        except (OSError, subprocess.SubprocessError):
            continue
        if probe.returncode == 0:
            shell = candidate
            break

    if shell is None:
        raise RCodeError("ไม่พบ PowerShell")

    with tempfile.TemporaryDirectory(prefix="thai-id-rcode-") as temp_dir:
        output = Path(temp_dir) / "rcode.xlsx"
        script = (
            "$ProgressPreference='SilentlyContinue'; "
            "$ErrorActionPreference='Stop'; "
            "Invoke-WebRequest -UseBasicParsing "
            "-Uri $env:RCODE_URL -OutFile $env:RCODE_OUT "
            "-TimeoutSec ([int]$env:RCODE_TIMEOUT)"
        )
        env = os.environ.copy()
        env["RCODE_URL"] = url
        env["RCODE_OUT"] = str(output)
        env["RCODE_TIMEOUT"] = str(max(1, int(timeout)))
        completed = subprocess.run(
            [shell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout + 10,
            env=env,
            check=False,
        )
        if completed.returncode != 0:
            message = (completed.stderr or completed.stdout or "PowerShell download failed").strip()
            raise RCodeError(message)
        content = output.read_bytes()
        _validate_xlsx_content(content)
        return content


def _download_with_curl(url: str, timeout: float) -> bytes:
    """Fallback อีกชั้นสำหรับ Windows 10/11 ที่มี curl.exe มาให้"""
    with tempfile.TemporaryDirectory(prefix="thai-id-rcode-") as temp_dir:
        output = Path(temp_dir) / "rcode.xlsx"
        command = [
            "curl.exe" if os.name == "nt" else "curl",
            "-L", "--fail", "--silent", "--show-error",
            "--max-time", str(max(1, int(timeout))),
            "--output", str(output),
            url,
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout + 10,
                check=False,
            )
        except (OSError, subprocess.SubprocessError) as error:
            raise RCodeError(f"เรียก curl ไม่สำเร็จ: {error}") from error
        if completed.returncode != 0:
            raise RCodeError((completed.stderr or completed.stdout or "curl download failed").strip())
        content = output.read_bytes()
        _validate_xlsx_content(content)
        return content


def _download_official_xlsx(
    request: urllib.request.Request,
    *,
    timeout: float,
) -> tuple[bytes, str, str | None, str | None]:
    """ลอง urllib ก่อน แล้วค่อย PowerShell/curl เมื่อการเชื่อมต่อปกติล้มเหลว"""
    errors: list[str] = []
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            content = _read_response_limited(response, MAX_DOWNLOAD_BYTES)
            _validate_xlsx_content(content)
            return (
                content,
                "urllib",
                response.headers.get("ETag"),
                response.headers.get("Last-Modified"),
            )
    except urllib.error.HTTPError as error:
        if error.code == 304:
            raise
        errors.append(f"urllib HTTP {error.code}: {error.reason}")
    except (urllib.error.URLError, TimeoutError, OSError, RCodeError, ValueError) as error:
        errors.append(f"urllib: {error}")

    # PowerShell ใช้ certificate store ของ Windows จึงช่วยกรณี Python SSL มีปัญหาได้
    if os.name == "nt":
        try:
            content = _download_with_powershell(SOURCE_XLSX_URL, timeout)
            return content, "powershell", None, None
        except (RCodeError, OSError, subprocess.SubprocessError) as error:
            errors.append(f"PowerShell: {error}")

    try:
        content = _download_with_curl(SOURCE_XLSX_URL, timeout)
        return content, "curl", None, None
    except (RCodeError, OSError, subprocess.SubprocessError) as error:
        errors.append(f"curl: {error}")

    raise RCodeError(" | ".join(errors))


def _read_response_limited(response: Any, maximum: int) -> bytes:
    content_length = response.headers.get("Content-Length")
    if content_length:
        try:
            if int(content_length) > maximum:
                raise RCodeError("ไฟล์ RCODE ใหญ่เกินขนาดที่กำหนด")
        except ValueError:
            pass

    chunks: list[bytes] = []
    total = 0
    while True:
        chunk = response.read(64 * 1024)
        if not chunk:
            break
        total += len(chunk)
        if total > maximum:
            raise RCodeError("ไฟล์ RCODE ใหญ่เกินขนาดที่กำหนด")
        chunks.append(chunk)
    return b"".join(chunks)


def _update_cache_error_metadata(
    database: dict[str, Any],
    path: Path,
    error: str,
) -> None:
    metadata = database.setdefault("metadata", {})
    metadata["last_checked_at"] = now_iso()
    metadata["last_error"] = error
    try:
        atomic_write_json(path, database)
    except OSError:
        # การอัปเดต metadata ไม่สำคัญเท่าการรักษา cache เดิมไว้
        pass


def import_local_xlsx(xlsx_path: Path, cache_path: Path) -> UpdateResult:
    checked_at = now_iso()
    try:
        content = xlsx_path.read_bytes()
        database = build_database(content)
        database["metadata"]["imported_from"] = str(xlsx_path.resolve())
        atomic_write_json(cache_path, database)
        return UpdateResult(
            attempted=True,
            status="imported",
            cache_path=str(cache_path),
            message=f"นำเข้า {xlsx_path.name} และสร้าง {cache_path.name} แล้ว",
            used_cache=True,
            downloaded=False,
            record_count=len(database["registries"]),
            data_as_of_th=database["metadata"].get("data_as_of_th"),
            checked_at=checked_at,
        )
    except (OSError, RCodeError) as error:
        return UpdateResult(
            attempted=True,
            status="import_failed",
            cache_path=str(cache_path),
            message=f"นำเข้า XLSX ไม่สำเร็จ: {error}",
            used_cache=cache_path.exists(),
            downloaded=False,
            checked_at=checked_at,
            error=str(error),
        )


def update_rcode_cache(
    cache_path: Path,
    *,
    timeout: float = DEFAULT_TIMEOUT,
    force: bool = False,
) -> UpdateResult:
    checked_at = now_iso()

    try:
        old_database = load_database(cache_path)
    except RCodeError:
        old_database = None
        try:
            broken_path = cache_path.with_suffix(
                cache_path.suffix + f".broken-{datetime.now():%Y%m%d-%H%M%S}"
            )
            cache_path.replace(broken_path)
        except OSError:
            pass

    old_metadata = (old_database or {}).get("metadata", {})
    headers = {
        "User-Agent": f"{APP_NAME}/{APP_VERSION} Python/{sys.version_info.major}",
        "Accept": (
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet,application/octet-stream,*/*"
        ),
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    }

    if not force:
        if old_metadata.get("http_etag"):
            headers["If-None-Match"] = str(old_metadata["http_etag"])
        if old_metadata.get("http_last_modified"):
            headers["If-Modified-Since"] = str(old_metadata["http_last_modified"])

    request = urllib.request.Request(SOURCE_XLSX_URL, headers=headers)

    try:
        content, method, etag, last_modified = _download_official_xlsx(
            request, timeout=timeout
        )
        new_hash = hashlib.sha256(content).hexdigest()
        old_hash = old_metadata.get("sha256")

        if old_database and old_hash == new_hash:
            metadata = old_database.setdefault("metadata", {})
            metadata["last_checked_at"] = checked_at
            metadata["last_successful_download_at"] = checked_at
            metadata["download_method"] = method
            metadata["http_etag"] = etag or metadata.get("http_etag")
            metadata["http_last_modified"] = (
                last_modified or metadata.get("http_last_modified")
            )
            metadata["last_error"] = None
            metadata["fallback"] = False
            atomic_write_json(cache_path, old_database)
            return UpdateResult(
                attempted=True,
                status="unchanged",
                cache_path=str(cache_path),
                message=(
                    f"ดาวน์โหลดไฟล์ทางการสำเร็จผ่าน {method} "
                    "แต่เนื้อหาเหมือนฐานเดิม จึงอัปเดตเฉพาะเวลาตรวจสอบ"
                ),
                used_cache=True,
                downloaded=True,
                record_count=len(old_database.get("registries", {})),
                data_as_of_th=metadata.get("data_as_of_th"),
                checked_at=checked_at,
            )

        database = build_database(
            content,
            http_etag=etag,
            http_last_modified=last_modified,
        )
        metadata = database["metadata"]
        metadata["last_successful_download_at"] = checked_at
        metadata["download_method"] = method
        metadata["fallback"] = False
        atomic_write_json(cache_path, database)
        return UpdateResult(
            attempted=True,
            status="updated" if old_database else "created",
            cache_path=str(cache_path),
            message=(
                f"ดาวน์โหลดผ่าน {method} และอัปเดต {cache_path.name} แล้ว"
                if old_database
                else f"ดาวน์โหลดผ่าน {method} และสร้าง {cache_path.name} แล้ว"
            ),
            used_cache=True,
            downloaded=True,
            record_count=len(database["registries"]),
            data_as_of_th=metadata.get("data_as_of_th"),
            checked_at=checked_at,
        )

    except urllib.error.HTTPError as error:
        if error.code == 304 and old_database:
            metadata = old_database.setdefault("metadata", {})
            metadata["last_checked_at"] = checked_at
            metadata["last_successful_download_at"] = checked_at
            metadata["last_error"] = None
            atomic_write_json(cache_path, old_database)
            return UpdateResult(
                attempted=True,
                status="not_modified",
                cache_path=str(cache_path),
                message="เซิร์ฟเวอร์ยืนยันว่าไฟล์ RCODE ยังไม่เปลี่ยน",
                used_cache=True,
                downloaded=False,
                record_count=len(old_database.get("registries", {})),
                data_as_of_th=metadata.get("data_as_of_th"),
                checked_at=checked_at,
            )
        failure: Exception = error
    except (
        urllib.error.URLError,
        TimeoutError,
        OSError,
        RCodeError,
        ValueError,
        subprocess.SubprocessError,
    ) as error:
        failure = error

    error_message = str(failure)

    if old_database:
        _update_cache_error_metadata(old_database, cache_path, error_message)
        metadata = old_database.get("metadata", {})
        return UpdateResult(
            attempted=True,
            status="cached_after_error",
            cache_path=str(cache_path),
            message="ติดต่อไฟล์ทางการไม่ได้ จึงใช้ rcode.json เดิมต่อ",
            used_cache=True,
            downloaded=False,
            record_count=len(old_database.get("registries", {})),
            data_as_of_th=metadata.get("data_as_of_th"),
            checked_at=checked_at,
            error=error_message,
        )

    fallback = create_fallback_database(error_message)
    try:
        atomic_write_json(cache_path, fallback)
        message = "ติดต่อไฟล์ทางการไม่ได้ จึงสร้าง rcode.json จาก snapshot ในโปรแกรม"
        used_cache = True
    except OSError as write_error:
        message = f"ดาวน์โหลดและสร้าง rcode.json ไม่สำเร็จ: {write_error}"
        used_cache = False
        error_message = f"{error_message}; write cache: {write_error}"

    return UpdateResult(
        attempted=True,
        status="fallback_created" if used_cache else "failed",
        cache_path=str(cache_path),
        message=message,
        used_cache=used_cache,
        downloaded=False,
        record_count=len(fallback.get("registries", {})),
        data_as_of_th=fallback.get("metadata", {}).get("data_as_of_th"),
        checked_at=checked_at,
        error=error_message,
    )


# ---------------------------------------------------------------------------
# Thai phone decoding
# ---------------------------------------------------------------------------


def _phone_compact(raw: str) -> tuple[str, bool]:
    if not isinstance(raw, str):
        raise ValueError("ข้อมูลต้องเป็นข้อความ")
    value = raw.strip().translate(THAI_DIGITS)
    if not value:
        raise ValueError("ไม่ได้กรอกเลขโทรศัพท์")

    # รูปเขียนสากลที่พบบ่อย เช่น +66 (0) 81 234 5678: (0) เป็นเพียง
    # notation ว่าเวลาโทรในประเทศให้ใส่ trunk prefix ไม่ใช่เลขที่ต้องกดหลัง +66
    value = re.sub(r"^\+66\s*\(\s*0\s*\)", "+66", value)

    # รองรับรูปแบบที่พบทั่วไป แต่ไม่ยอมรับ extension/ตัวอักษรเพื่อเลี่ยงการเดาผิด
    if re.search(r"[A-Za-zก-๙]", value):
        raise ValueError("เบอร์โทรต้องไม่มีตัวอักษรหรือ extension")
    if not re.fullmatch(r"[0-9+().\s\-–—]+", value):
        raise ValueError("พบอักขระที่ไม่รองรับในเบอร์โทร")
    if value.count("+") > 1 or ("+" in value and not value.lstrip().startswith("+")):
        raise ValueError("เครื่องหมาย + ใช้ได้เฉพาะด้านหน้า country code")

    compact = re.sub(r"[().\s\-–—]", "", value)
    formatted = compact != value
    return compact, formatted


def normalize_phone(raw: str) -> tuple[str, str]:
    """คืน (เลขรูปแบบภายในประเทศ, input_format). ไม่ยืนยันว่าเลขถูกจัดสรรจริง."""
    compact, had_separators = _phone_compact(raw)

    # เลขสั้นและเลขบริการพิเศษบางกลุ่มไม่มี trunk prefix 0
    if re.fullmatch(r"[0-9]{3,4}", compact):
        return compact, "short_code"
    if re.fullmatch(r"(?:1401|1800|1900)[0-9]{6}", compact):
        return compact, "special_service"

    if compact.startswith("+66"):
        nsn = compact[3:]
        if nsn.startswith("0"):
            raise ValueError("รูปแบบ +66 ต้องตัดเลข 0 ด้านหน้าของเบอร์ไทยออก เช่น +66812345678")
        if not re.fullmatch(r"[0-9]{8,9}", nsn):
            raise ValueError("หลัง +66 ต้องมี National Significant Number 8 หรือ 9 หลัก")
        return "0" + nsn, "international_formatted" if had_separators else "e164"

    if compact.startswith("0066"):
        nsn = compact[4:]
        if nsn.startswith("0"):
            raise ValueError("รูปแบบ 0066 ต้องตัดเลข 0 ด้านหน้าของเบอร์ไทยออก")
        if not re.fullmatch(r"[0-9]{8,9}", nsn):
            raise ValueError("หลัง 0066 ต้องมี National Significant Number 8 หรือ 9 หลัก")
        return "0" + nsn, "international_00"

    # ยอมรับ country code แบบไม่มี + เพื่อความสะดวก เช่น 66812345678
    if compact.startswith("66") and len(compact) in {10, 11}:
        nsn = compact[2:]
        if nsn.startswith("0"):
            raise ValueError("หลัง country code 66 ต้องไม่มี trunk prefix 0")
        return "0" + nsn, "country_code_no_plus"

    if compact.startswith("+"):
        raise ValueError("รองรับเฉพาะเลขประเทศไทย country code +66")

    if not compact.isdigit():
        raise ValueError("เบอร์โทรต้องประกอบด้วยตัวเลข")

    if len(compact) not in {9, 10}:
        raise ValueError(
            "เลขโทรศัพท์ไทยทั่วไปควรมี 9 หลัก (ประจำที่) หรือ 10 หลัก (มือถือ/VoIP) "
            "เมื่อเขียนแบบภายในประเทศ"
        )
    if not compact.startswith("0"):
        raise ValueError("รูปแบบภายในประเทศต้องขึ้นต้นด้วย trunk prefix 0")
    return compact, "domestic_formatted" if had_separators else "domestic_plain"


def format_phone(national: str) -> str:
    if len(national) == 10 and national.startswith("0"):
        return f"{national[:3]}-{national[3:6]}-{national[6:]}"
    if len(national) == 9 and national.startswith("02"):
        return f"{national[:2]}-{national[2:5]}-{national[5:]}"
    if len(national) == 9 and national.startswith("0"):
        return f"{national[:3]}-{national[3:6]}-{national[6:]}"
    if len(national) == 3:
        return national
    if len(national) == 4:
        return national
    if len(national) == 10 and national[:4] in {"1401", "1800", "1900"}:
        return f"{national[:4]}-{national[4:7]}-{national[7:]}"
    return national


def mask_phone(national: str) -> str:
    formatted = format_phone(national)
    if len(national) == 10 and national.startswith("0"):
        return f"{national[:3]}-XXX-{national[-4:]}"
    if len(national) == 9 and national.startswith("02"):
        return f"{national[:2]}-XXX-{national[-4:]}"
    if len(national) == 9 and national.startswith("0"):
        return f"{national[:3]}-XXX-{national[-3:]}"
    if len(national) > 4:
        return national[:4] + "-XXX-" + national[-3:]
    return formatted


def lookup_historical_operator(national: str) -> dict[str, Any] | None:
    for prefix in sorted(MOBILE_PREFIX_OPERATOR_HINTS, key=len, reverse=True):
        if national.startswith(prefix):
            item = dict(MOBILE_PREFIX_OPERATOR_HINTS[prefix])
            item.update(
                {
                    "matched_prefix": prefix,
                    "meaning": "historical/original-allocation hint only",
                    "current_carrier": None,
                    "current_carrier_reason": "ยืนยันไม่ได้จาก prefix เพราะมี Mobile Number Portability (MNP)",
                    "source_id": "wikipedia_th_phone_prefix_snapshot",
                }
            )
            return item
    return None


def lookup_prefix_history(national: str) -> dict[str, Any] | None:
    for prefix in sorted(PREFIX_HISTORY, key=len, reverse=True):
        if national.startswith(prefix):
            item = dict(PREFIX_HISTORY[prefix])
            item["matched_prefix"] = prefix
            return item
    return None


def detect_input_kind(raw: str) -> str:
    """แยก ID/phone แบบไม่ใช้ฐานข้อมูลภายนอก."""
    if not isinstance(raw, str):
        return "unknown"
    value = raw.strip().translate(THAI_DIGITS)
    if not value:
        return "unknown"

    if FORMATTED_PATTERN.fullmatch(value) or PLAIN_PATTERN.fullmatch(value):
        return "thai_id"

    # ถ้ามีตัวเลขรวม 13 หลักและไม่ได้ขึ้นต้นด้วย 0 ให้ถือว่า user ตั้งใจกรอกบัตร
    digits_only = re.sub(r"\D", "", value)
    if len(digits_only) == 13 and not digits_only.startswith("0"):
        return "thai_id"

    try:
        normalize_phone(value)
    except ValueError:
        return "unknown"
    return "thai_phone"


def _fixed_area_for_number(national: str) -> tuple[str | None, dict[str, Any] | None]:
    if national.startswith("02"):
        return "02", FIXED_AREA_CODES.get("02")
    if len(national) >= 3:
        code = national[:3]
        return code, FIXED_AREA_CODES.get(code)
    return None, None


def decode_phone(raw: str, *, mask: bool = False) -> PhoneDecodeResult:
    result = PhoneDecodeResult(raw_input=raw)
    result.source_notes = [dict(item) for item in PHONE_SOURCES]

    try:
        national, input_format = normalize_phone(raw)
    except ValueError as error:
        result.errors.append(str(error))
        return result

    result.normalized_national = national
    result.formatted_national = format_phone(national)
    result.displayed_number = mask_phone(national) if mask else result.formatted_national
    result.input_format = input_format
    result.format_valid = True

    # Short numbers (3/4 digits)
    if len(national) in {3, 4}:
        result.number_type = "short_code"
        result.numbering_category = "เลขหมายโทรศัพท์แบบสั้น"
        result.prefix = national
        result.special_service = SHORT_NUMBER_HINTS.get(national)
        result.structurally_valid = True
        result.assignment_verified = result.special_service is not None
        if result.special_service is None:
            result.warnings.append(
                "รูปแบบเป็นเลขสั้นที่เป็นไปได้ แต่โปรแกรมไม่ได้ฝังฐานเลขสั้น 4 หลักทั้งหมด; "
                "ควรตรวจการจัดสรรล่าสุดกับ NBTC"
            )
        return result

    # Toll-free/premium/other special leading-group numbers
    if len(national) == 10 and national[:4] in {"1401", "1800", "1900"}:
        group = national[:4]
        labels = {
            "1401": "บริการพิเศษ/โทรคมนาคมรูปแบบพิเศษ",
            "1800": "Toll Free Service",
            "1900": "Premium-Based Service",
        }
        result.number_type = "special_service"
        result.numbering_category = labels[group]
        result.prefix = group
        result.subscriber_number = national[4:]
        result.structurally_valid = True
        result.assignment_verified = False
        result.warnings.append("ตรวจได้เฉพาะโครงสร้างกลุ่มเลข ไม่ได้ยืนยันว่าเลขปลายทางถูกจัดสรรหรือเปิดใช้งาน")
        return result

    result.trunk_prefix = "0"
    result.national_significant_number = national[1:]
    result.e164 = "+66" + national[1:]

    # Fixed-line: ปัจจุบันเลขภายในประเทศ 9 หลัก (รวม trunk 0)
    if len(national) == 9:
        area_code, area = _fixed_area_for_number(national)
        result.area_code = area_code
        result.area = dict(area) if area else None
        result.prefix = area_code
        if area_code == "02":
            result.subscriber_number = national[2:]
        else:
            result.subscriber_number = national[3:]

        if area is None:
            result.number_type = "unknown_fixed_like"
            result.numbering_category = "รูปแบบความยาวเหมือนเลขประจำที่ แต่ไม่พบรหัสพื้นที่ใน snapshot"
            result.errors.append("ไม่พบรหัสพื้นที่ที่ตรงกับแผนเลขหมายที่โปรแกรมรองรับ")
            return result

        result.number_type = "fixed_line"
        result.numbering_category = "โทรศัพท์ประจำที่ (geographic fixed line)"
        result.structurally_valid = True
        result.assignment_verified = False
        if not area.get("active", False):
            result.warnings.append("รหัสนี้เป็นกลุ่มสำรองในแผนเลขหมาย จึงไม่ควรตีความว่าเลขรายนี้ถูกจัดสรรใช้งานจริง")
        return result

    # 10-digit public/technical ranges
    prefix3 = national[:3]
    prefix2 = national[:2]
    result.prefix = prefix3
    result.subscriber_number = national[3:]
    result.prefix_history = lookup_prefix_history(national)
    result.historical_operator_hint = lookup_historical_operator(national)

    if prefix3 == "060":
        result.number_type = "voip"
        result.numbering_category = "Voice over Internet Protocol (VoIP) / non-geographic"
        result.structurally_valid = True
    elif prefix3 in {"061", "062", "063", "064", "065", "066"}:
        result.number_type = "mobile"
        result.numbering_category = "โทรศัพท์เคลื่อนที่ / non-geographic"
        result.structurally_valid = True
    elif prefix3 == "067":
        result.number_type = "iot_reserved_family"
        result.numbering_category = "หมวด 067 ใช้กับเลขหมาย IoT ตามแผนเลขหมาย ไม่ใช่รูปแบบมือถือ 10 หลักทั่วไป"
        result.structurally_valid = False
        result.errors.append("067 ไม่ควรถูกตีความเป็นเบอร์มือถือ 10 หลักทั่วไป")
    elif prefix3 == "068":
        result.number_type = "technical_routing"
        result.numbering_category = "หมวดเลขด้านเทคนิค/ประวัติ VoIP; ไม่ควรยืนยันเป็นเบอร์มือถือจากรูปแบบอย่างเดียว"
        result.structurally_valid = False
        result.errors.append("068 มีการใช้งานด้าน routing/เลขทางเทคนิคในแผนปัจจุบัน จึงไม่ยืนยันเป็นเลขผู้ใช้ปลายทาง")
    elif prefix3 == "069":
        result.number_type = "technical_routing"
        result.numbering_category = "Routing Number สำหรับบริการโทรศัพท์ระหว่างประเทศ"
        result.structurally_valid = False
        result.errors.append("069 เป็นหมวดด้าน routing ตามแผนเลขหมาย ไม่ใช่เบอร์มือถือผู้ใช้ทั่วไป")
    elif prefix2 in {"08", "09"}:
        result.number_type = "mobile"
        result.numbering_category = "โทรศัพท์เคลื่อนที่ / non-geographic"
        result.structurally_valid = True
        if prefix3 == "089":
            result.warnings.append(
                "แหล่งแผนเลขหมายและตาราง historical บางชุดจัดหมวด 089 ต่างกัน; "
                "โปรแกรมจึงถือเป็นเลขมือถือที่พบใช้งานจริง แต่ไม่ใช้หมวดนี้ยืนยันค่าย/เทคโนโลยี"
            )
    elif prefix2 in {"04", "05", "07"}:
        result.number_type = "planned_mobile_range"
        result.numbering_category = "หมวดที่แผนเลขหมายรองรับสำหรับ mobile ในโครงสร้างใหม่/สำรอง"
        result.structurally_valid = True
        result.warnings.append(
            "ผ่านโครงสร้างตามหมวดเลข แต่ไม่ได้หมายความว่าช่วงย่อยนี้ถูกจัดสรรหรือเปิดใช้งานในปัจจุบัน"
        )
    else:
        result.number_type = "unknown_10_digit"
        result.numbering_category = "ไม่ตรงหมวดเลขโทรศัพท์ไทยทั่วไปที่โปรแกรมยืนยันได้"
        result.errors.append("prefix ไม่ตรงหมวด mobile/VoIP/special ที่รู้จัก")

    result.assignment_verified = False
    result.current_operator_verified = False
    result.issue_year = None
    if result.structurally_valid:
        result.warnings.append(
            "ค่ายปัจจุบันยืนยันไม่ได้จาก prefix เพราะผู้ใช้สามารถย้ายค่ายเบอร์เดิม (MNP) ได้"
        )
        result.warnings.append(
            "ปีที่ออกเลข/ปีเปิดซิมของเลขรายนี้ไม่ได้ถูก encode อยู่ในตัวเลข; แสดงได้เพียงประวัติของหมวด prefix"
        )
        result.warnings.append(
            "การผ่าน format/numbering-plan ไม่ได้ยืนยันว่าเลขนี้ถูกจัดสรร มีผู้ใช้ หรือยังเปิดใช้งานอยู่"
        )
    return result


# ---------------------------------------------------------------------------
# ID decoding
# ---------------------------------------------------------------------------


def normalize_id(raw: str, *, lenient: bool = False) -> tuple[str, str]:
    if not isinstance(raw, str):
        raise ValueError("ข้อมูลต้องเป็นข้อความ")

    value = raw.strip().translate(THAI_DIGITS)
    if not value:
        raise ValueError("ไม่ได้กรอกเลขประจำตัว")

    if PLAIN_PATTERN.fullmatch(value):
        return value, "plain"

    if FORMATTED_PATTERN.fullmatch(value):
        return value.replace("-", ""), "formatted"

    if lenient:
        compact = re.sub(r"[\s\-–—]", "", value)
        if PLAIN_PATTERN.fullmatch(compact):
            return compact, "lenient"

    raise ValueError(
        "รูปแบบไม่ถูกต้อง ต้องเป็น 1234567890123 "
        "หรือ 1-2345-67890-12-3"
    )


def format_id(person_id: str) -> str:
    return (
        f"{person_id[0]}-{person_id[1:5]}-{person_id[5:10]}-"
        f"{person_id[10:12]}-{person_id[12]}"
    )


def mask_id(person_id: str) -> str:
    return f"{person_id[0]}-{person_id[1:5]}-XXXXX-XX-{person_id[12]}"


def calculate_checksum(person_id: str) -> ChecksumResult:
    if not PLAIN_PATTERN.fullmatch(person_id):
        raise ValueError("checksum ต้องรับเลข 13 หลักที่ normalize แล้ว")

    products: list[dict[str, int]] = []
    total = 0
    for index, character in enumerate(person_id[:12]):
        digit = int(character)
        weight = 13 - index
        product = digit * weight
        total += product
        products.append(
            {
                "position": index + 1,
                "digit": digit,
                "weight": weight,
                "product": product,
            }
        )

    remainder = total % 11
    expected = (11 - remainder) % 10
    actual = int(person_id[12])
    return ChecksumResult(
        actual=actual,
        expected=expected,
        valid=actual == expected,
        total=total,
        remainder=remainder,
        products=products,
    )


def serial_meaning(person_type_code: int) -> dict[str, str]:
    """
    หลัก 6-10 และ 11-12 มีความหมายเป็นกลุ่ม/ลำดับ หรือเล่ม/ใบที่
    แล้วแต่กระบวนการทางทะเบียน จึงไม่ควรระบุว่าเป็นสูติบัตรเสมอไป
    """
    if person_type_code in {1, 2}:
        return {
            "group_label": "กลุ่ม/เล่มทะเบียน",
            "sequence_label": "ลำดับ/ใบที่",
            "note": (
                "สำหรับการแจ้งเกิดอาจสัมพันธ์กับเล่มและใบสูติบัตร "
                "แต่ไม่มีฐานสาธารณะสำหรับถอดกลับเป็นวันเกิดหรือบุคคล"
            ),
        }

    return {
        "group_label": "กลุ่มทะเบียน",
        "sequence_label": "ลำดับในกลุ่ม",
        "note": (
            "ความหมายย่อยขึ้นกับประเภทและกระบวนการทางทะเบียน "
            "ไม่สามารถใช้หาอายุ วันเกิด หรือความสัมพันธ์ได้"
        ),
    }


def decode_id(
    raw: str,
    *,
    database: dict[str, Any] | None,
    cache_path: Path,
    lenient: bool = False,
    mask: bool = False,
) -> DecodeResult:
    result = DecodeResult(raw_input=raw)
    result.database = database_summary(database, cache_path)

    try:
        person_id, input_format = normalize_id(raw, lenient=lenient)
    except ValueError as error:
        result.errors.append(str(error))
        return result

    result.normalized_id = person_id
    result.formatted_id = format_id(person_id)
    result.displayed_id = mask_id(person_id) if mask else result.formatted_id
    result.input_format = input_format
    result.format_valid = True

    person_type_code = int(person_id[0])
    person_type = PERSON_TYPES.get(person_type_code)
    result.person_type_known = person_type is not None
    result.person_type = {
        "code": person_type_code,
        "recognized": result.person_type_known,
        "title": person_type["title"] if person_type else "ไม่พบประเภทบุคคล",
        "detail": (
            person_type["detail"]
            if person_type
            else "หลักแรกไม่อยู่ในประเภท 0-8 ที่รองรับ"
        ),
        "current_status_verified": False,
    }

    registry_code = person_id[1:5]
    province_code = registry_code[:2]
    local_code = registry_code[2:]
    registries = (database or {}).get("registries", {})
    provinces = (database or {}).get("provinces", {})
    registry_record = registries.get(registry_code)

    province_record = provinces.get(province_code, {})
    province_name = (
        (registry_record or {}).get("province_name_th")
        or province_record.get("name_th")
        or PROVINCES_TH.get(province_code)
    )

    result.registry_known = registry_record is not None
    result.registry_active = (
        registry_record.get("active") if registry_record else None
    )
    result.reference_match = result.registry_known

    result.registry = {
        "code": registry_code,
        "known": result.registry_known,
        "active": result.registry_active,
        "name_th": (registry_record or {}).get("name_th"),
        "name_en": (registry_record or {}).get("name_en"),
        "office_type": (registry_record or {}).get("office_type"),
        "office_type_th": (registry_record or {}).get("office_type_th"),
        "locality_name_th": (registry_record or {}).get("locality_name_th"),
        "province_code": province_code,
        "province_name_th": province_name,
        "province_name_en": (registry_record or {}).get("province_name_en"),
        "local_code": local_code,
        "discontinued_be": (registry_record or {}).get("discontinued_be"),
        "discontinued_iso": (registry_record or {}).get("discontinued_iso"),
        "meaning": (
            "สำนักทะเบียนที่ใช้จัดสรรเลขในตอนแรก "
            "ไม่ใช่ที่อยู่ปัจจุบันและไม่ยืนยันสถานที่เกิด"
        ),
    }

    serial_info = serial_meaning(person_type_code)
    group_or_book = person_id[5:10]
    sequence_or_sheet = person_id[10:12]
    result.serial = {
        "combined": person_id[5:12],
        "group_or_book": group_or_book,
        "sequence_or_sheet": sequence_or_sheet,
        "group_label": serial_info["group_label"],
        "sequence_label": serial_info["sequence_label"],
        "note": serial_info["note"],
        "reverse_lookup_available": False,
    }

    checksum = calculate_checksum(person_id)
    result.checksum = asdict(checksum)
    result.checksum_valid = checksum.valid

    registry_code_sane = registry_code != "0000"
    result.structurally_valid = (
        result.format_valid
        and result.person_type_known
        and registry_code_sane
        and result.checksum_valid
    )

    if not result.person_type_known:
        result.errors.append(
            f"ประเภทบุคคลหลักแรกเป็น {person_type_code} ซึ่งไม่อยู่ในช่วง 0-8"
        )

    if not registry_code_sane:
        result.errors.append("รหัสสำนักทะเบียนเป็น 0000 ซึ่งผิดปกติ")

    if not result.checksum_valid:
        result.errors.append(
            "checksum ไม่ตรง "
            f"(หลักสุดท้ายเป็น {checksum.actual}, ควรเป็น {checksum.expected})"
        )

    if registry_code_sane and not result.registry_known:
        if province_name:
            result.warnings.append(
                f"ทราบเพียงจังหวัด {province_name} แต่ไม่พบ RCODE {registry_code} "
                "ในฐานที่โหลดอยู่"
            )
        else:
            result.warnings.append(
                f"ไม่พบ RCODE {registry_code} และไม่รู้จักรหัสจังหวัด {province_code}"
            )

    if result.registry_known and result.registry_active is False:
        discontinued = (
            result.registry.get("discontinued_be")
            or result.registry.get("discontinued_iso")
            or "ไม่ระบุวันที่"
        )
        result.warnings.append(
            f"RCODE {registry_code} ถูกระบุว่าเลิกใช้/จำหน่ายแล้ว ({discontinued}) "
            "แต่เลขที่ออกในอดีตอาจยังคงใช้รหัสเดิม"
        )

    if group_or_book == "00000":
        result.warnings.append("ค่ากลุ่ม/เล่มเป็น 00000 ซึ่งควรตรวจสอบเพิ่มเติม")

    if sequence_or_sheet == "00":
        result.warnings.append("ค่าลำดับ/ใบที่เป็น 00 ซึ่งควรตรวจสอบเพิ่มเติม")

    result.warnings.append(
        "ผ่าน checksum หมายถึงตัวเลขสัมพันธ์ตามสูตรเท่านั้น "
        "ไม่ยืนยันว่ามีบุคคลนี้จริงในฐานข้อมูลรัฐ"
    )

    return result


# ---------------------------------------------------------------------------
# Lookup / search / compare
# ---------------------------------------------------------------------------


def lookup_rcode(database: dict[str, Any] | None, code: str) -> dict[str, Any]:
    normalized = code.strip().translate(THAI_DIGITS)
    if not RCODE_PATTERN.fullmatch(normalized):
        raise ValueError("RCODE ต้องเป็นเลข 4 หลัก เช่น 3097")

    registries = (database or {}).get("registries", {})
    record = registries.get(normalized)
    return {
        "code": normalized,
        "found": record is not None,
        "record": record,
        "fallback_province": PROVINCES_TH.get(normalized[:2]),
    }


def search_rcode(
    database: dict[str, Any] | None,
    query: str,
    *,
    limit: int = 30,
) -> list[dict[str, Any]]:
    needle = normalize_search_text(query.translate(THAI_DIGITS))
    if not needle:
        return []

    registries = (database or {}).get("registries", {})
    matches: list[tuple[int, str, dict[str, Any]]] = []

    for code, record in registries.items():
        fields = [
            code,
            str(record.get("name_th") or ""),
            str(record.get("name_en") or ""),
            str(record.get("locality_name_th") or ""),
            str(record.get("province_name_th") or ""),
            str(record.get("office_type_th") or ""),
        ]
        normalized_fields = [normalize_search_text(field) for field in fields]

        score = 0
        if code == needle:
            score += 100
        elif code.startswith(needle):
            score += 60

        if any(field == needle for field in normalized_fields[1:]):
            score += 50
        elif any(field.startswith(needle) for field in normalized_fields[1:]):
            score += 30
        elif any(needle in field for field in normalized_fields[1:]):
            score += 10

        if score:
            matches.append((score, code, record))

    matches.sort(key=lambda item: (-item[0], item[1]))
    return [
        {"code": code, **record}
        for _, code, record in matches[: max(1, limit)]
    ]


def compare_ids(
    raw_a: str,
    raw_b: str,
    *,
    database: dict[str, Any] | None,
    cache_path: Path,
    lenient: bool = False,
    mask: bool = False,
) -> dict[str, Any]:
    first = decode_id(
        raw_a,
        database=database,
        cache_path=cache_path,
        lenient=lenient,
        mask=mask,
    )
    second = decode_id(
        raw_b,
        database=database,
        cache_path=cache_path,
        lenient=lenient,
        mask=mask,
    )

    comparison: dict[str, Any] = {
        "first": asdict(first),
        "second": asdict(second),
        "comparable": first.format_valid and second.format_valid,
        "same_person_type": None,
        "same_registry": None,
        "same_group_or_book": None,
        "sequence_distance": None,
        "warning": (
            "ผลนี้เปรียบเทียบเพียงโครงสร้างตัวเลข "
            "ไม่ยืนยันว่าเป็นญาติ อยู่บ้านเดียวกัน หรือเกิดในช่วงเดียวกัน"
        ),
    }

    if comparison["comparable"]:
        comparison["same_person_type"] = (
            first.person_type["code"] == second.person_type["code"]
        )
        comparison["same_registry"] = (
            first.registry["code"] == second.registry["code"]
        )
        comparison["same_group_or_book"] = (
            first.serial["group_or_book"] == second.serial["group_or_book"]
        )
        if (
            comparison["same_person_type"]
            and comparison["same_registry"]
            and comparison["same_group_or_book"]
        ):
            comparison["sequence_distance"] = abs(
                int(first.serial["sequence_or_sheet"])
                - int(second.serial["sequence_or_sheet"])
            )

    return comparison


# ---------------------------------------------------------------------------
# Human-readable output
# ---------------------------------------------------------------------------


def update_status_line(update: UpdateResult) -> str:
    if update.status in {"updated", "created", "imported"}:
        state = "อัปเดตแล้ว"
    elif update.status in {"unchanged", "not_modified"}:
        state = "เป็นข้อมูลล่าสุดจากไฟล์ต้นทาง"
    elif update.status == "cached_after_error":
        state = "อัปเดตไม่ได้ ใช้ cache เดิม"
    elif update.status == "fallback_created":
        state = "ใช้ snapshot ในตัวโปรแกรม"
    elif update.status == "skipped":
        state = "ข้ามการอัปเดต"
    else:
        state = "ไม่พร้อมใช้งาน"

    extras: list[str] = []
    if update.record_count:
        extras.append(f"{update.record_count:,} รหัส")
    if update.data_as_of_th:
        extras.append(f"ข้อมูล ณ {update.data_as_of_th}")

    suffix = f" ({', '.join(extras)})" if extras else ""
    return f"RCODE : {state}{suffix}"


def print_update_status(update: UpdateResult, *, verbose: bool = False) -> None:
    print(update_status_line(update))
    print(f"JSON  : {update.cache_path}")
    print(f"ผลอัปเดต : {update.message}")
    if update.error:
        print(f"สาเหตุที่ใช้ออนไลน์ไม่ได้ : {update.error}")
    if verbose:
        print(f"สถานะภายใน : {update.status}")
        print(f"ดาวน์โหลดสำเร็จ : {update.downloaded}")
        print(f"เวลาตรวจสอบ : {update.checked_at}")


def print_decode(result: DecodeResult, *, verbose: bool = False) -> None:
    print()
    print("ผลตรวจสอบเลขประจำตัว 13 หลัก")
    print("=" * 62)

    if not result.format_valid:
        print("สถานะ           : ไม่ผ่าน")
        print(f"สาเหตุ           : {result.errors[0] if result.errors else 'ไม่ทราบ'}")
        return

    if result.structurally_valid and result.registry_known:
        status = "ผ่านโครงสร้าง และพบรหัสสำนักทะเบียน"
    elif result.structurally_valid:
        status = "ผ่านโครงสร้าง แต่ไม่พบรหัสสำนักทะเบียนในฐาน"
    else:
        status = "ไม่ผ่านการตรวจสอบเชิงโครงสร้าง"

    print(f"เลข              : {result.displayed_id}")
    print(f"สถานะ            : {status}")
    print(f"Checksum         : {'ถูกต้อง' if result.checksum_valid else 'ไม่ถูกต้อง'}")

    if result.errors:
        print("\nข้อผิดพลาด")
        for error in result.errors:
            print(f"- {error}")

    print("\nข้อมูลที่อ่านได้")
    print(f"ประเภทตอนออกเลข : {result.person_type['title']}")

    registry = result.registry
    if registry["known"]:
        print(f"สำนักทะเบียน     : {registry['name_th']}")
        print(f"ประเภทสำนักงาน   : {registry['office_type_th']}")
        if registry.get("locality_name_th"):
            print(f"ชื่อพื้นที่        : {registry['locality_name_th']}")
        print(f"จังหวัด           : {registry['province_name_th'] or 'ไม่ทราบ'}")
        active_text = "ใช้งานอยู่ในชุดข้อมูล" if registry["active"] else "ยกเลิก/จำหน่ายแล้ว"
        print(f"รหัส RCODE       : {registry['code']} ({active_text})")
        if not registry["active"] and registry.get("discontinued_be"):
            print(f"วันที่ยกเลิก      : {registry['discontinued_be']} พ.ศ.")
    else:
        print(f"รหัส RCODE       : {registry['code']} (ไม่พบในฐาน)")
        print(f"จังหวัดจากรหัส   : {registry['province_name_th'] or 'ไม่ทราบ'}")

    serial = result.serial
    print(f"{serial['group_label']:<17}: {serial['group_or_book']}")
    print(f"{serial['sequence_label']:<17}: {serial['sequence_or_sheet']}")

    print("\nข้อจำกัด")
    print(f"- {registry['meaning']}")
    print(f"- {serial['note']}")
    print("- เลขนี้ไม่บอกชื่อ วันเกิด อายุ เพศ ที่อยู่ปัจจุบัน หรือเจ้าของเลข")
    print("- โปรแกรมไม่สามารถยืนยันว่าเลขนี้มีอยู่จริงในฐานข้อมูลรัฐ")

    other_warnings = [
        warning
        for warning in result.warnings
        if not warning.startswith("ผ่าน checksum หมายถึง")
    ]
    if other_warnings:
        print("\nคำเตือน")
        for warning in other_warnings:
            print(f"- {warning}")

    if verbose:
        checksum = result.checksum
        formula = " + ".join(
            f"{item['digit']}x{item['weight']}"
            for item in checksum["products"]
        )
        print("\nรายละเอียดสำหรับนักพัฒนา")
        print("-" * 62)
        print(f"Normalized ID    : {result.normalized_id}")
        print(f"Input format     : {result.input_format}")
        print(f"Person type      : {result.person_type['code']}")
        print(f"Type detail      : {result.person_type['detail']}")
        print(f"Province code    : {registry['province_code']}")
        print(f"Local code       : {registry['local_code']}")
        print(f"Serial combined  : {serial['combined']}")
        print(f"Checksum formula : {formula}")
        print(f"Checksum total   : {checksum['total']}")
        print(f"Total mod 11     : {checksum['remainder']}")
        print(f"Actual / expected: {checksum['actual']} / {checksum['expected']}")
        print(f"Structure valid  : {result.structurally_valid}")
        print(f"Registry known   : {result.registry_known}")
        print("Existence checked: False")


def print_phone_decode(result: PhoneDecodeResult, *, verbose: bool = False) -> None:
    print()
    print("ผลตรวจสอบเลขโทรศัพท์ไทย")
    print("=" * 72)

    if not result.format_valid:
        print("สถานะ            : ไม่ผ่านรูปแบบ")
        print(f"สาเหตุ            : {result.errors[0] if result.errors else 'ไม่ทราบ'}")
        return

    print(f"เลข              : {result.displayed_number}")
    if result.e164:
        print(f"E.164            : {result.e164}")
    print(f"ประเภท           : {result.numbering_category or result.number_type or 'ไม่ทราบ'}")
    print(f"โครงสร้าง        : {'ผ่าน' if result.structurally_valid else 'ไม่ผ่าน/ยืนยันไม่ได้'}")

    if result.area_code:
        print(f"รหัสพื้นที่        : {result.area_code}")
        if result.area:
            print(f"ภูมิภาค           : {result.area.get('region_th') or '-'}")
            provinces = result.area.get("provinces_th") or []
            if provinces:
                print(f"จังหวัดที่ครอบคลุม : {', '.join(provinces)}")
            if result.area.get("note"):
                print(f"หมายเหตุพื้นที่    : {result.area['note']}")

    if result.prefix and not result.area_code:
        print(f"Prefix           : {result.prefix}")
    if result.subscriber_number:
        print(f"ส่วน Subscriber  : {result.subscriber_number}")

    if result.special_service:
        print("\nข้อมูลเลขสั้นจาก กสทช.")
        print(f"หน่วยงาน          : {result.special_service.get('organization') or '-'}")
        print(f"ภารกิจ/บริการ     : {result.special_service.get('purpose') or '-'}")

    print("\nข้อมูลค่าย/การจัดสรร")
    hint = result.historical_operator_hint
    if hint:
        print(f"ค่ายจาก prefix เดิม: {hint.get('operator')} ({hint.get('confidence')})")
        print(f"จับคู่จาก prefix   : {hint.get('matched_prefix')}")
        print("ค่ายปัจจุบัน       : ยืนยันไม่ได้จากเลขอย่างเดียว (MNP)")
    elif result.number_type in {"mobile", "voip", "planned_mobile_range"}:
        print("ค่ายจาก prefix เดิม: ไม่มีข้อมูลที่เชื่อถือได้พอใน snapshot นี้")
        print("ค่ายปัจจุบัน       : ยืนยันไม่ได้จากเลขอย่างเดียว (MNP)")
    else:
        print("ค่ายปัจจุบัน       : ไม่เกี่ยวข้อง/ยืนยันไม่ได้จาก prefix")

    print("ปีออกเลขรายนี้     : ถอดจากตัวเลขไม่ได้")
    if result.prefix_history:
        print(f"ประวัติ prefix     : {result.prefix_history.get('event')}")
        print(f"หมายเหตุปี         : {result.prefix_history.get('important')}")

    if result.errors:
        print("\nข้อผิดพลาด/ข้อจำกัดโครงสร้าง")
        for error in result.errors:
            print(f"- {error}")

    if result.warnings:
        print("\nคำเตือน")
        for warning in result.warnings:
            print(f"- {warning}")

    if verbose:
        print("\nรายละเอียดสำหรับนักพัฒนา")
        print("-" * 72)
        print(f"Normalized        : {result.normalized_national}")
        print(f"Input format      : {result.input_format}")
        print(f"NSN               : {result.national_significant_number}")
        print(f"Number type       : {result.number_type}")
        print(f"Assignment checked: {result.assignment_verified}")
        print(f"Current carrier   : verified={result.current_operator_verified}")
        print(f"Source snapshot   : {result.source_snapshot_date}")


def print_lookup(result: dict[str, Any]) -> None:
    print()
    print("ผลค้นหา RCODE")
    print("=" * 62)
    print(f"รหัส             : {result['code']}")

    if not result["found"]:
        print("สถานะ           : ไม่พบในฐาน RCODE")
        if result.get("fallback_province"):
            print(f"จังหวัดจาก prefix: {result['fallback_province']}")
        return

    record = result["record"]
    print(f"ชื่อ             : {record.get('name_th') or '-'}")
    print(f"ชื่ออังกฤษ       : {record.get('name_en') or '-'}")
    print(f"ประเภท           : {record.get('office_type_th') or '-'}")
    print(f"พื้นที่           : {record.get('locality_name_th') or '-'}")
    print(f"จังหวัด          : {record.get('province_name_th') or '-'}")
    print(
        "สถานะ           : "
        + ("ใช้งานอยู่ในชุดข้อมูล" if record.get("active") else "ยกเลิก/จำหน่ายแล้ว")
    )
    if record.get("discontinued_be"):
        print(f"วันที่ยกเลิก      : {record['discontinued_be']} พ.ศ.")


def print_search(results: list[dict[str, Any]], query: str) -> None:
    print()
    print(f"ผลค้นหา RCODE: {query}")
    print("=" * 78)
    if not results:
        print("ไม่พบข้อมูล")
        return

    for record in results:
        status = "ใช้งาน" if record.get("active") else "ยกเลิก"
        print(
            f"{record['code']} | {record.get('name_th') or '-'} | "
            f"{record.get('province_name_th') or '-'} | {status}"
        )


def print_compare(comparison: dict[str, Any]) -> None:
    print()
    print("เปรียบเทียบโครงสร้างเลขสองชุด")
    print("=" * 62)
    first = comparison["first"]
    second = comparison["second"]

    print(f"เลข A : {first.get('displayed_id') or first.get('raw_input')}")
    print(f"เลข B : {second.get('displayed_id') or second.get('raw_input')}")

    if not comparison["comparable"]:
        print("สถานะ : เปรียบเทียบไม่ได้ เพราะมีเลขอย่างน้อยหนึ่งชุดผิดรูปแบบ")
        return

    def yes_no(value: bool | None) -> str:
        if value is None:
            return "ไม่ทราบ"
        return "ใช่" if value else "ไม่ใช่"

    print(f"ประเภทเดียวกัน       : {yes_no(comparison['same_person_type'])}")
    print(f"สำนักทะเบียนเดียวกัน : {yes_no(comparison['same_registry'])}")
    print(f"กลุ่ม/เล่มเดียวกัน   : {yes_no(comparison['same_group_or_book'])}")
    distance = comparison.get("sequence_distance")
    print(f"ระยะห่างของลำดับ     : {distance if distance is not None else 'คำนวณไม่ได้'}")
    print(f"\nหมายเหตุ: {comparison['warning']}")


# ---------------------------------------------------------------------------
# Self-test and CLI
# ---------------------------------------------------------------------------


def make_valid_test_id(first_12: str) -> str:
    if not re.fullmatch(r"[0-9]{12}", first_12):
        raise ValueError("ต้องเป็นเลข 12 หลัก")
    temporary = first_12 + "0"
    return first_12 + str(calculate_checksum(temporary).expected)


def run_self_test() -> tuple[bool, list[str]]:
    messages: list[str] = []

    test_id = make_valid_test_id("140980012886")
    if test_id != "1409800128861":
        messages.append(f"checksum sample ผิด: ได้ {test_id}")

    try:
        normalized, input_format = normalize_id("๑-๔๐๙๘-๐๐๑๒๘-๘๖-๑")
        if normalized != "1409800128861" or input_format != "formatted":
            messages.append("normalize เลขไทยผิด")
    except ValueError as error:
        messages.append(f"normalize เลขไทยเกิด error: {error}")

    # Phone normalization / classification tests
    phone_cases = [
        ("081-234-5678", "0812345678", "thai_phone"),
        ("+66 81 234 5678", "0812345678", "thai_phone"),
        ("02-123-4567", "021234567", "thai_phone"),
        ("044-213-456", "044213456", "thai_phone"),
    ]
    for raw_phone, expected_phone, expected_kind in phone_cases:
        try:
            normalized_phone, _fmt = normalize_phone(raw_phone)
        except ValueError as error:
            messages.append(f"normalize phone {raw_phone} error: {error}")
            continue
        if normalized_phone != expected_phone:
            messages.append(f"normalize phone {raw_phone} ผิด: {normalized_phone}")
        if detect_input_kind(raw_phone) != expected_kind:
            messages.append(f"auto classify {raw_phone} ผิด")

    mobile_test = decode_phone("081-234-5678")
    if not mobile_test.structurally_valid or mobile_test.number_type != "mobile":
        messages.append("decode mobile 081 ผิด")

    fixed_test = decode_phone("044-213-456")
    if not fixed_test.structurally_valid or fixed_test.area_code != "044":
        messages.append("decode fixed 044 ผิด")

    if detect_input_kind(test_id) != "thai_id":
        messages.append("auto classify Thai ID ผิด")

    try:
        content = Path(__file__).resolve().with_name("rcode.xlsx").read_bytes()
    except OSError:
        content = b""

    if content:
        try:
            database = build_database(content)
            if "3097" not in database["registries"]:
                messages.append("ไม่พบ RCODE 3097 ใน workbook ทดสอบ")
        except RCodeError as error:
            messages.append(f"parse workbook ทดสอบไม่สำเร็จ: {error}")

    return not messages, messages


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "รับเลขเดียวแล้วแยกอัตโนมัติว่าเป็นเลขบัตรประชาชนไทย 13 หลัก "
            "หรือเลขโทรศัพท์ไทย พร้อม decode โครงสร้างที่อ้างอิงได้"
        )
    )

    parser.add_argument(
        "value",
        nargs="?",
        help=("เลขบัตร 13 หลัก หรือเบอร์โทรไทย เช่น 0812345678, "
              "+66812345678, 02-123-4567"),
    )
    parser.add_argument(
        "--input-type",
        choices=("auto", "id", "phone"),
        default="auto",
        help="บังคับชนิด input; ค่าเริ่มต้น auto ให้โปรแกรมแยกบัตร/เบอร์เอง",
    )
    parser.add_argument(
        "--cache",
        type=Path,
        default=DEFAULT_CACHE,
        help=f"ตำแหน่ง rcode.json (ค่าเริ่มต้น: {DEFAULT_CACHE.name})",
    )
    parser.add_argument(
        "--no-update",
        action="store_true",
        help="ไม่เช็กอัปเดต RCODE ในรอบนี้",
    )
    parser.add_argument(
        "--force-update",
        action="store_true",
        help="ดาวน์โหลด RCODE ใหม่โดยไม่ใช้ conditional request",
    )
    parser.add_argument(
        "--update-only",
        action="store_true",
        help="อัปเดตฐาน RCODE แล้วจบโปรแกรม",
    )
    parser.add_argument(
        "--import-xlsx",
        type=Path,
        metavar="FILE",
        help="สร้าง rcode.json จาก rcode.xlsx ที่ดาวน์โหลดไว้เอง",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=DEFAULT_TIMEOUT,
        help=f"timeout การดาวน์โหลดเป็นวินาที (ค่าเริ่มต้น {DEFAULT_TIMEOUT:g})",
    )
    parser.add_argument(
        "--lookup",
        metavar="RCODE",
        help="ค้นรหัสสำนักทะเบียน 4 หลัก เช่น 3097",
    )
    parser.add_argument(
        "--find",
        metavar="TEXT",
        help="ค้น RCODE จากชื่อ เช่น บัวใหญ่ หรือ Bua Yai",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=30,
        help="จำนวนผลลัพธ์สูงสุดของ --find",
    )
    parser.add_argument(
        "--compare",
        nargs=2,
        metavar=("ID_A", "ID_B"),
        help="เปรียบเทียบโครงสร้างเลขสองชุด",
    )
    parser.add_argument(
        "--lenient",
        action="store_true",
        help="ยอมรับช่องว่างหรือขีดหลายรูปแบบแล้ว normalize ให้อัตโนมัติ",
    )
    parser.add_argument(
        "--mask",
        action="store_true",
        help="ซ่อนบางหลักของเลขบัตร/เบอร์โทรใน output สำหรับ screenshot/log",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="แสดงรายละเอียดสำหรับนักพัฒนา (รวม checksum สำหรับเลขบัตร)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="แสดงผลเป็น JSON เท่านั้น",
    )
    parser.add_argument(
        "--source",
        action="store_true",
        help="แสดงแหล่งข้อมูล RCODE และแหล่งอ้างอิงเลขโทรศัพท์ แล้วจบโปรแกรม",
    )
    parser.add_argument(
        "--self-test",
        action="store_true",
        help="ทดสอบฟังก์ชันหลักภายในโปรแกรม",
    )
    return parser


def main() -> int:
    configure_console()
    parser = build_parser()
    args = parser.parse_args()
    cache_path = args.cache.expanduser().resolve()

    if args.source:
        payload = {
            "rcode": {
                "source_page": SOURCE_PAGE_URL,
                "source_xlsx": SOURCE_XLSX_URL,
                "cache": str(cache_path),
            },
            "phone": {
                "reference_checked": PHONE_REFERENCE_CHECKED,
                "sources": PHONE_SOURCES,
            },
        }
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print(f"หน้าข้อมูล RCODE : {SOURCE_PAGE_URL}")
            print(f"ไฟล์ RCODE       : {SOURCE_XLSX_URL}")
            print(f"ไฟล์ cache       : {cache_path}")
            print("\nแหล่งอ้างอิงเลขโทรศัพท์:")
            for source in PHONE_SOURCES:
                print(f"- [{source['authority']}] {source['title']}: {source['url']}")
        return 0

    if args.self_test:
        passed, messages = run_self_test()
        payload = {"passed": passed, "messages": messages}
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print("SELF-TEST: " + ("ผ่าน" if passed else "ไม่ผ่าน"))
            for message in messages:
                print(f"- {message}")
        return 0 if passed else 1

    # โหมด input ปกติ: อ่านค่าก่อน เพื่อให้เบอร์โทรไม่ต้องเสียเวลาวิ่งอัปเดต RCODE
    raw_value: str | None = None
    ordinary_input_mode = not any(
        [args.import_xlsx, args.update_only, args.lookup, args.find, args.compare]
    )
    if ordinary_input_mode:
        raw_value = args.value
        if raw_value is None:
            try:
                raw_value = input("กรอกเลขบัตรประชาชน 13 หลัก หรือเบอร์โทรศัพท์ไทย: ")
            except (EOFError, KeyboardInterrupt):
                print("\nยกเลิก")
                return 130

        if args.input_type == "phone":
            input_kind = "thai_phone"
        elif args.input_type == "id":
            input_kind = "thai_id"
        else:
            input_kind = detect_input_kind(raw_value)

        # ถ้า auto แล้วไม่เหมือน ID ให้ phone decoder อธิบาย format error โดยตรง
        if input_kind == "thai_phone" or (input_kind == "unknown" and args.input_type == "auto"):
            phone_result = decode_phone(raw_value, mask=args.mask)
            payload = {
                "input_kind": "thai_phone" if phone_result.format_valid else "unknown",
                "result": asdict(phone_result),
            }
            if args.json:
                print(json.dumps(payload, ensure_ascii=False, indent=2))
            else:
                print_phone_decode(phone_result, verbose=args.verbose)
            return 0 if phone_result.structurally_valid else 1

    if args.import_xlsx:
        update = import_local_xlsx(args.import_xlsx.expanduser(), cache_path)
    elif args.no_update:
        try:
            existing = load_database(cache_path)
        except RCodeError as error:
            existing = None
            skipped_error = str(error)
        else:
            skipped_error = None

        metadata = (existing or {}).get("metadata", {})
        update = UpdateResult(
            attempted=False,
            status="skipped",
            cache_path=str(cache_path),
            message="ข้ามการอัปเดตตาม --no-update",
            used_cache=existing is not None,
            downloaded=False,
            record_count=len((existing or {}).get("registries", {})),
            data_as_of_th=metadata.get("data_as_of_th"),
            checked_at=now_iso(),
            error=skipped_error,
        )
    else:
        update = update_rcode_cache(
            cache_path,
            timeout=max(1.0, args.timeout),
            force=args.force_update,
        )

    try:
        database = load_database(cache_path)
    except RCodeError as error:
        database = create_fallback_database(str(error))

    if args.update_only:
        payload = {"update": asdict(update), "database": database_summary(database, cache_path)}
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print_update_status(update, verbose=args.verbose)
        return 0 if update.used_cache else 1

    if args.lookup:
        try:
            lookup = lookup_rcode(database, args.lookup)
        except ValueError as error:
            parser.error(str(error))

        payload = {
            "update": asdict(update),
            "database": database_summary(database, cache_path),
            "lookup": lookup,
        }
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print_update_status(update, verbose=args.verbose)
            print_lookup(lookup)
        return 0 if lookup["found"] else 1

    if args.find:
        results = search_rcode(database, args.find, limit=max(1, args.limit))
        payload = {
            "update": asdict(update),
            "database": database_summary(database, cache_path),
            "query": args.find,
            "results": results,
        }
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print_update_status(update, verbose=args.verbose)
            print_search(results, args.find)
        return 0 if results else 1

    if args.compare:
        comparison = compare_ids(
            args.compare[0],
            args.compare[1],
            database=database,
            cache_path=cache_path,
            lenient=args.lenient,
            mask=args.mask,
        )
        payload = {
            "update": asdict(update),
            "database": database_summary(database, cache_path),
            "comparison": comparison,
        }
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print_update_status(update, verbose=args.verbose)
            print_compare(comparison)
        return 0 if comparison["comparable"] else 1

    # ถึงจุดนี้เป็น ID mode (หรือ --input-type id)
    if raw_value is None:
        raw_value = args.value
    if raw_value is None:
        try:
            raw_value = input("กรอกเลขประจำตัว 13 หลัก: ")
        except (EOFError, KeyboardInterrupt):
            print("\nยกเลิก")
            return 130

    result = decode_id(
        raw_value,
        database=database,
        cache_path=cache_path,
        lenient=args.lenient,
        mask=args.mask,
    )

    payload = {
        "input_kind": "thai_id",
        "update": asdict(update),
        "result": asdict(result),
    }
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print_update_status(update, verbose=args.verbose)
        print_decode(result, verbose=args.verbose)

    return 0 if result.structurally_valid else 1

# ---------------------------------------------------------------------------
# Terminal TUI frontend
# ---------------------------------------------------------------------------

def _tui_enable_ansi() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        try:
            os.system("")
        except OSError:
            pass


_TUI_RESET = "\x1b[0m"
_TUI_WHITE = "\x1b[97m"
_TUI_GRAY = "\x1b[90m"

_TUI_LOGO = r"""                               .
                              ..
                             ..
                            ...
                          ....       .
                         .. ..      ..
                        .. ..     ...
                       .. ..     ...
                    ....  ...    ...
      .............           ..........
           ..........   ............
                   ..  ..      ...
                 . .. .. ..................
               .. .. .. ......................
                  . .     .......
                 ....     ......
                 ..      ......     ID / Number Lookup
                ..      .... .      By Zipher@Nickqme
               ..       ...
               .       ...
                       ..
                      ..
                      ."""


def _tui_enter() -> None:
    sys.stdout.write(_TUI_RESET + "\x1b[?1049h\x1b[2J\x1b[H")
    sys.stdout.flush()


def _tui_leave() -> None:
    sys.stdout.write(_TUI_RESET + "\x1b[?1049l")
    sys.stdout.flush()


def _tui_begin_frame() -> None:
    sys.stdout.write("\x1b[2J\x1b[H")


def _tui_value(value: Any) -> bool:
    return value is not None and value != "" and value != [] and value != {}


def _tui_field(label: str, value: Any) -> str | None:
    if not _tui_value(value):
        return None
    return f"{_TUI_GRAY}{label}{_TUI_RESET}  {_TUI_WHITE}{value}{_TUI_RESET}"


def _tui_id_data_lines(result: DecodeResult) -> list[str]:
    lines: list[str] = []

    def add(label: str, value: Any) -> None:
        line = _tui_field(label, value)
        if line:
            lines.append(line)

    add("ID", result.displayed_id or result.formatted_id or result.normalized_id)

    person = result.person_type or {}
    if person.get("recognized"):
        code = person.get("code")
        title = person.get("title")
        add("PERSON TYPE", f"{code}  {title}" if title else code)
        add("TYPE DETAIL", person.get("detail"))

    registry = result.registry or {}
    add("RCODE", registry.get("code"))
    if registry.get("known"):
        add("REGISTRY", registry.get("name_th"))
        add("OFFICE TYPE", registry.get("office_type_th"))
        add("LOCALITY", registry.get("locality_name_th"))
        add("PROVINCE", registry.get("province_name_th"))
        add("DISCONTINUED", registry.get("discontinued_be") or registry.get("discontinued_iso"))
    else:
        add("PROVINCE", registry.get("province_name_th"))

    serial = result.serial or {}
    add(serial.get("group_label") or "GROUP", serial.get("group_or_book"))
    add(serial.get("sequence_label") or "SEQUENCE", serial.get("sequence_or_sheet"))
    return lines


def _tui_phone_data_lines(result: PhoneDecodeResult) -> list[str]:
    lines: list[str] = []

    def add(label: str, value: Any) -> None:
        line = _tui_field(label, value)
        if line:
            lines.append(line)

    add("NUMBER", result.displayed_number or result.formatted_national or result.normalized_national)
    add("TYPE", result.numbering_category or result.number_type)
    add("E.164", result.e164)
    add("PREFIX", result.prefix)
    add("SUBSCRIBER", result.subscriber_number)

    if result.area_code:
        add("AREA CODE", result.area_code)

    area = result.area or {}
    add("REGION", area.get("region_th"))
    provinces = area.get("provinces_th") or []
    if provinces:
        add("PROVINCES", ", ".join(str(x) for x in provinces if x))

    special = result.special_service or {}
    add("ORGANIZATION", special.get("organization"))
    add("SERVICE", special.get("purpose"))

    operator = result.historical_operator_hint or {}
    add("ORIGINAL CARRIER", operator.get("operator"))
    add("MATCHED PREFIX", operator.get("matched_prefix"))

    history = result.prefix_history or {}
    add("PREFIX HISTORY", history.get("event"))
    return lines


def _tui_rcode_data_lines(lookup: dict[str, Any]) -> list[str]:
    lines: list[str] = []

    def add(label: str, value: Any) -> None:
        line = _tui_field(label, value)
        if line:
            lines.append(line)

    add("RCODE", lookup.get("code"))
    record = lookup.get("record") or {}
    if lookup.get("found"):
        add("REGISTRY", record.get("name_th"))
        add("OFFICE TYPE", record.get("office_type_th"))
        add("LOCALITY", record.get("locality_name_th"))
        add("PROVINCE", record.get("province_name_th"))
        add("DISCONTINUED", record.get("discontinued_be") or record.get("discontinued_iso"))
    else:
        add("PROVINCE", lookup.get("fallback_province"))
    return lines


def _tui_notes(kind: str, result: Any) -> list[str]:
    notes: list[str] = []

    def push(value: Any) -> None:
        if value is None:
            return
        value = str(value).strip()
        if value and value not in notes:
            notes.append(value)

    if kind == "rcode":
        if not result.get("found"):
            push(f"ไม่พบ RCODE {result.get('code') or ''} ในฐานข้อมูล")
        return notes

    if kind == "phone":
        for error in result.errors:
            push(error)
        for warning in result.warnings:
            push(warning)
        area = result.area or {}
        push(area.get("note"))
        return notes

    if kind == "id":
        for error in result.errors:
            push(error)
        for warning in result.warnings:
            if warning and not str(warning).startswith("ผ่าน checksum หมายถึง"):
                push(warning)
        registry = result.registry or {}
        if result.format_valid and not registry.get("known"):
            push(registry.get("meaning"))
        return notes

    push(result)
    return notes


def _tui_data_lines(kind: str, result: Any) -> list[str]:
    if kind == "id":
        return _tui_id_data_lines(result)
    if kind == "phone":
        return _tui_phone_data_lines(result)
    if kind == "rcode":
        return _tui_rcode_data_lines(result)
    return []


def _tui_render(*, last_result: tuple[str, Any] | None, query_count: int) -> None:
    _tui_begin_frame()
    sys.stdout.write(_TUI_WHITE + _TUI_LOGO + _TUI_RESET + "\n")

    if last_result is not None:
        kind, result = last_result
        sys.stdout.write(f"\n{_TUI_GRAY}-> [{query_count}] ------------------{_TUI_RESET}\n")
        for line in _tui_data_lines(kind, result):
            sys.stdout.write(line + "\n")
        for note in _tui_notes(kind, result):
            line = _tui_field("NOTE", note)
            if line:
                sys.stdout.write(line + "\n")
        sys.stdout.write(f"{_TUI_GRAY}-> ----------------------{_TUI_RESET}\n\n")

    sys.stdout.write(f"{_TUI_WHITE}INPUT:{_TUI_RESET}\n")
    sys.stdout.flush()


def _tui_windows_readline(prompt: str, history: list[str]) -> str:
    import msvcrt

    buffer: list[str] = []
    cursor = 0
    history_index = len(history)
    draft = ""
    previous_render_len = 0

    def redraw() -> None:
        nonlocal previous_render_len
        text = "".join(buffer)
        visible = prompt + text
        padding = " " * max(0, previous_render_len - len(visible))
        sys.stdout.write("\r" + visible + padding)
        move_left = len(text) - cursor + len(padding)
        if move_left:
            sys.stdout.write(f"\x1b[{move_left}D")
        sys.stdout.flush()
        previous_render_len = len(visible)

    redraw()
    while True:
        char = msvcrt.getwch()

        if char == "\r":
            return "".join(buffer)
        if char == "\x03":
            raise KeyboardInterrupt
        if char == "\x08":
            if cursor > 0:
                del buffer[cursor - 1]
                cursor -= 1
                redraw()
            continue

        if char in ("\x00", "\xe0"):
            key = msvcrt.getwch()
            if key == "H":
                if history:
                    if history_index == len(history):
                        draft = "".join(buffer)
                    history_index = max(0, history_index - 1)
                    buffer[:] = list(history[history_index])
                    cursor = len(buffer)
                    redraw()
            elif key == "P":
                if history:
                    history_index = min(len(history), history_index + 1)
                    value = draft if history_index == len(history) else history[history_index]
                    buffer[:] = list(value)
                    cursor = len(buffer)
                    redraw()
            elif key == "K":
                if cursor > 0:
                    cursor -= 1
                    redraw()
            elif key == "M":
                if cursor < len(buffer):
                    cursor += 1
                    redraw()
            elif key == "G":
                cursor = 0
                redraw()
            elif key == "O":
                cursor = len(buffer)
                redraw()
            elif key == "S":
                if cursor < len(buffer):
                    del buffer[cursor]
                    redraw()
            continue

        if ord(char) < 32:
            continue
        buffer.insert(cursor, char)
        cursor += 1
        redraw()


def _tui_readline(prompt: str, history: list[str]) -> str:
    if os.name == "nt":
        return _tui_windows_readline(prompt, history)

    try:
        import readline
    except ImportError:
        return input(prompt)

    try:
        readline.clear_history()
        for item in history:
            readline.add_history(item)
    except Exception:
        pass
    return input(prompt)


def build_tui_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="TUI - Thai ID / Phone / Number Lookup")
    parser.add_argument(
        "--cache",
        type=Path,
        default=DEFAULT_CACHE,
        help=f"ตำแหน่ง rcode.json (ค่าเริ่มต้น: {DEFAULT_CACHE.name})",
    )
    parser.add_argument("--no-update", action="store_true")
    parser.add_argument("--force-update", action="store_true")
    parser.add_argument(
        "--timeout",
        type=float,
        default=DEFAULT_TIMEOUT,
    )
    parser.add_argument("--mask", action="store_true")
    return parser


def _tui_load_database(
    args: argparse.Namespace,
    cache_path: Path,
) -> tuple[dict[str, Any] | None, UpdateResult]:
    if args.no_update:
        try:
            database = load_database(cache_path)
            error = None
        except RCodeError as exc:
            database = None
            error = str(exc)

        if database is None:
            database = create_fallback_database(error or "ไม่มี rcode.json")
            try:
                atomic_write_json(cache_path, database)
            except OSError:
                pass

        metadata = (database or {}).get("metadata", {})
        return database, UpdateResult(
            attempted=False,
            status="skipped",
            cache_path=str(cache_path),
            message="ข้ามการอัปเดตตาม --no-update",
            used_cache=database is not None,
            downloaded=False,
            record_count=len((database or {}).get("registries", {})),
            data_as_of_th=metadata.get("data_as_of_th"),
            checked_at=now_iso(),
            error=error,
        )

    update = update_rcode_cache(
        cache_path,
        timeout=max(1.0, args.timeout),
        force=args.force_update,
    )
    try:
        database = load_database(cache_path)
    except RCodeError as exc:
        database = create_fallback_database(str(exc))
    return database, update


def _tui_decode_input(
    raw: str,
    *,
    database: dict[str, Any] | None,
    cache_path: Path,
    mask: bool,
) -> tuple[str, Any]:
    translated = raw.strip().translate(THAI_DIGITS)

    # Keep direct RCODE lookup from the previous TUI without breaking known
    # phone short-codes from the original source.
    if RCODE_PATTERN.fullmatch(translated) and translated not in SHORT_NUMBER_HINTS:
        try:
            lookup = lookup_rcode(database, translated)
        except ValueError:
            lookup = None
        if lookup and lookup.get("found"):
            return "rcode", lookup

    kind = detect_input_kind(raw)
    if kind == "thai_id":
        return "id", decode_id(
            raw,
            database=database,
            cache_path=cache_path,
            lenient=True,
            mask=mask,
        )

    # Same auto behavior as the uploaded original: anything that is not
    # classified as a Thai ID is handed to the Thai phone decoder.
    return "phone", decode_phone(raw, mask=mask)


def tui_main() -> int:
    configure_console()
    _tui_enable_ansi()
    args = build_tui_parser().parse_args()
    cache_path = args.cache.expanduser().resolve()

    history: list[str] = []
    query_count = 0
    last_result: tuple[str, Any] | None = None

    _tui_enter()
    try:
        _tui_render(last_result=None, query_count=0)
        database, _update = _tui_load_database(args, cache_path)

        while True:
            _tui_render(last_result=last_result, query_count=query_count)
            try:
                raw = _tui_readline("> ", history).strip()
            except (EOFError, KeyboardInterrupt):
                return 0

            if not raw:
                continue
            if raw.casefold() in {"q", "quit", "exit"}:
                return 0

            if not history or history[-1] != raw:
                history.append(raw)

            query_count += 1
            last_result = _tui_decode_input(
                raw,
                database=database,
                cache_path=cache_path,
                mask=args.mask,
            )
    finally:
        _tui_leave()


if __name__ == "__main__":
    raise SystemExit(tui_main())

