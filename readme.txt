유의할 사항: Group1.vbg를 수행한 후 실행하면 오류가 발생하므로
인스톨 한 후 POCKEKQUIZ.vbp를 수행하여 개발해야 한다. 2009.12.28

mysql> select * from tg02 ;
+--------+--------+--------------------------------------+
| grp    | code   | codeNM                               |
+--------+--------+--------------------------------------+
| ?      | 100000 | 리소스분류                           |
| ?      | 200000 | 테이블목록표                         |
| ?      | 210000 | 테이블명세표                         |
| ?      | 220000 | mode                                 |
| 100000 | @      | 텍스트                               |
| 100000 | ⓗ     | html                                 |
| 100000 | ⓘ     | image(gif)                           |
| 100000 | ⓜ     | multi media(video)                   |
| 100000 | ⓢ     | sound(mp3)                           |
| 200000 | TG01   | 학습상태                             |
| 200000 | TG02   | 공통코드                             |
| 200000 | TG03   | 최근문제풀이내역                     |
| 200000 | TH01   | 사용자힌트                           |
| 200000 | TP01   | 포켓퀴즈생성정보                     |
| 200000 | TP02   | 포켓퀴즈마스타                       |
| 200000 | TP03   | 포켓퀴즈학습상태                     |
| 200000 | TQ01   | 퀴즈01-4지선다과목                   |
| 200000 | TQ02   | 퀴즈02-단답식                        |
| 200000 | TQ04   | 중학영어                             |
| 200000 | TS01   | 과목관리                             |
| 200000 | TS02   | 사용자별과목                         |
| 200000 | TTP1   | 금언                                 |
| 200000 | TU01   | 사용자기본정보                       |
| 200000 | TU02   | 사용자학습상태                       |
| 200000 | TU03   | 사용자학습책갈피                     |
| 200000 | VQ01   | 퀴즈 뷰                              |
| 220000 | 1      | 4지선다 1 혹은A가 답                 |
| 220000 | 2      | mode2 는 o x 퀴즈 보기에 답이 없음   |
| 220000 | 4      | 4지선다형으로 ABCD 1234등을 답으로함 |
| 220000 | 5      | 5지선다형                            |
| 300000 | mode   | 단순사실은mode=2                     |
+--------+--------+--------------------------------------+
	VQ01_LOG 추가해야 함.

mysql> select * from tg01 limit 2;
+-----------+--------+----------+----------+---------+------------+-------+-------+---------+-------------+----------------+----------+------+------+------+------+------+------+------+------+
| gijunilja | userid | subj     | totalsec | new_cnt | review_cnt | o_cnt | x_cnt | chk_cnt | backNew_cnt | backReview_cnt | next_cnt | yyyy | q    | m    | y  | d    | w    | ww   | h    |
+-----------+--------+----------+----------+---------+------------+-------+-------+---------+-------------+----------------+----------+------+------+------+------+------+------+------+------+
| 20040622  | 일어   | 히라가나 |       27 |       0 |          0 |     0 |0 |       0 |        NULL |              0 |        0 | 2004 |    2 |    6 |  174 |   22 |    3 |   26 |    2 |
| 20040624  | 일어   | 히라가나 |     1053 |      24 |         87 |    56 |    24 |       0 |        NULL |              5 |       99 | 2004 |    2 |    6 |  176 |   24 |    5 |   26 |    2 |
+-----------+--------+----------+----------+---------+------------+-------+-------+---------+-------------+----------------+----------+------+------+------+------+------+------+------+------+
2 rows in set (0.00 sec)

mysql> select * from tg02 limit 2;
+-----+--------+--------------+
| grp | code   | codeNM       |
+-----+--------+--------------+
| ?   | 100000 | 리소스분류   |
| ?   | 200000 | 테이블목록표 |
+-----+--------+--------------+
2 rows in set (0.00 sec)

mysql> select * from tg03 limit 2;
Empty set (0.00 sec)

mysql> select * from th01 limit 2;
+--------+------+--------+---------------------------------------+--------+
| subj   | seq  | userid | hint                                  | bshare |
+--------+------+--------+---------------------------------------+--------+
| 1998AA |   41 | 시사   | 시스템분석가는 요구사항분석하는 사람~ | 1      |
| 1998AA |   42 | 시사   |                                       | 1      |
+--------+------+--------+---------------------------------------+--------+
2 rows in set (0.00 sec)

mysql> select * from tp01 limit 2;
+------+--------------+--------------------+------+------+
| seq  | pocketnm     | cond               | xm   | chkm |
+------+--------------+--------------------+------+------+
|   10 | CISA1998기출 | subj='1998AA'|tq01 |    0 |    0 |
|   11 | CISA1999기출 | subj='1999AA'|tq01 |    0 |    0 |
+------+--------------+--------------------+------+------+
*************************** 218. row ***************************
     seq: 1366
pocketnm: 신약:66
    cond: subj='신약성경66-요한계시록'|(select * from tq07 where cat like '_신:66%')
      xm: 0
    chkm: 0
*************************** 219. row ***************************
     seq: 808
pocketnm: 토익990
    cond: subj='토익990'|tq14
      xm: 0
    chkm: 0
2 rows in set (0.00 sec)

mysql> select * from tp02 limit 2;
+------+--------------+-------+--------+--------+------------+-------+------+------+-------+-------------+----------+----------+
| code | pocketnm     | chasu | userid | hidden | create_ymd | pcode | xm   | chkm | image | selectimage | from_ymd | to_ymd   |
+------+--------------+-------+--------+--------+------------+-------+------+------+-------+-------------+----------+----------+
|  187 | ▲ 히라가나2 |     0 | 일어   |      1 | 20040822   |     0 |    0 | 0 |     1 |           2 | 20040823 | 20040922 |
|  183 | ◆ 히라가나  |     0 | 일어   |      1 | 20040822   |     0 |    0 | 0 |     1 |           2 | 20040823 | 20040830 |
+------+--------------+-------+--------+--------+------------+-------+------+------+-------+-------------+----------+----------+
2 rows in set (0.00 sec)

mysql> select * from tp03 limit 2;
+---------------+-------+--------+------+----------+------+------+------+------+------------+
| pocketnm      | chasu | userid | num  | subj     | seq  | o    | x    | chk  | update_ymd |
+---------------+-------+--------+------+----------+------+------+------+------+------------+
| 17 회차 11-09 |     1 | 신승희 |   48 | 공인1-05 |    1 |    0 |    0 |    0 | 20041024   |
| 17 회차 11-09 |     1 | 신승희 |   49 | 공인1-01 |    5 |    0 |    0 |    0 | 20041024   |
+---------------+-------+--------+------+----------+------+------+------+------+------------+


mysql> select * from ts01 limit 2;
+--------+---------------+
| subj   | subjnm        |
+--------+---------------+
| 1998AA | CISA-1998문제 |
| 1999AA | CISA-1999문제 |
+--------+---------------+
2 rows in set (0.00 sec)

mysql> select * from ts02 limit 2;
+--------+--------+----------+----------+
| subj   | userid | startymd | endymd   |
+--------+--------+----------+----------+
| 1998AA | aaa    | 20000101 | 21001231 |
| 1998AA | 일어   | 20000101 | 21001231 |
+--------+--------+----------+----------+
2 rows in set (0.00 sec)

mysql> select * from tt01 limit 2;
+---------+------------+---------+
| userid  | subj       | sortkey |
+---------+------------+---------+
| 박규선2 | 중학기초   |       1 |
| 신승희  | 조리사0507 |       1 |
+---------+------------+---------+
2 rows in set (0.03 sec)

mysql> select * from tt02 limit 2;
+---------+----------+------+--------+
| userid  | subj     | seq  | nansu  |
+---------+----------+------+--------+
| 박규선2 | 중학기초 |    1 | 479134 |
| 박규선2 | 중학기초 |    2 | 821537 |
+---------+----------+------+--------+
2 rows in set (0.03 sec)

mysql> select * from tt03 limit 2;
+--------+------------+------+---------+-------+----------+----------+------+-------+
| userid | subj       | seq  | sortkey | chasu | fromilja | toilja   | num  | nansu |
+--------+------------+------+---------+-------+----------+----------+------+-------+
| 신승희 | 조리사0507 |    1 |    NULL |     1 | 20051120 | 20051120 |    1 |  NULL |
| 신승희 | 조리사0507 |    2 |    NULL |     1 | 20051120 | 20051120 |    2 |  NULL |
+--------+------------+------+---------+-------+----------+----------+------+-------+
2 rows in set (0.05 sec)

mysql> select * from ttp1 limit 2;
+------+--------------------------------------------------------+---------------------------+------+
| code | title                                                  | author            | cat  |
+------+--------------------------------------------------------+---------------------------+------+
|    1 | 많은 일을 하고자 하면 지금 당장 한 가지 일을 시작하라. | 로스차일드 (미국, 은행가) | NULL |
|    2 | 모든 일을 참되고 실속이 있도록 애써 행하라. (務實力行) | 안창호 (1878-1938)        | NULL |
+------+--------------------------------------------------------+---------------------------+------+
2 rows in set (0.02 sec)

mysql> select * from tu01 limit 2;
+--------+----------+----------+----------+
| userid | userpass | con_ymd  | con_time |
+--------+----------+----------+----------+
| 박규선 |          | 20100105 | 125726   |
| 부동산 |          | 20041127 | 132224   |
+--------+----------+----------+----------+
2 rows in set (0.00 sec)

mysql> select * from tu02 limit 2;
+--------+------+--------+------+------+------+------------+-------------+---------+
| subj   | seq  | userid | o    | x    | chk  | update_ymd | reserve_ymd | gangyek |
+--------+------+--------+------+------+------+------------+-------------+---------+
| 1998AA |    1 | 박규선 |    3 |    1 |    0 | 20040612   | 99999999    | 0 |
| 1998AA |    1 | 부동산 |    0 |    0 |    0 | 20040718   | 99999999    | 0 |
+--------+------+--------+------+------+------+------------+-------------+---------+
2 rows in set (0.08 sec)

mysql> select * from tu03 limit 2;
+--------------+-------+--------+---------+--------+----+
| pocketnm     | chasu | userid | lastnew | hangsu | od |
+--------------+-------+--------+---------+--------+----+
| ▲ 히라가나2 |     0 | 일어   |       1 |      1 | 10 |
| ◆ 히라가나  |     0 | 일어   |       1 |      1 | 10 |
+--------------+-------+--------+---------+--------+----+
2 rows in set (0.02 sec)

mysql> select * from ver limit 2;
+--------------+------+-----------+------------+
| DBID         | VER  | UPDATEYMD | UPDATETIME |
+--------------+------+-----------+------------+
| KR0001000001 |    5 | 20040922  | 000001     |
+--------------+------+-----------+------------+
1 row in set (0.01 sec)

mysql> select * from vq01 limit 2;
+--------+------+-------+----------------------------------------------------------------------------------------------+----------------------------------------+------------------------------------------------+-------------------------------------------------+-------------------------------------------------------+------+------------------------------------------------------------------------------------------------------------------------------------+------+-------+------+-----------+-------------+
| subj   | seq  | cat   | quiz                                       | a| b                                              | c                   | d                                                     | e  | hint                                                       | ans  | resid | mode | updateymd | updatechasu |
+--------+------+-------+----------------------------------------------------------------------------------------------+----------------------------------------+------------------------------------------------+-------------------------------------------------+-------------------------------------------------------+------+------------------------------------------------------------------------------------------------------------------------------------+------+-------+------+-----------+-------------+
| 1998AA |   60 | D3-14 | 기계어 형태의 업무프로그램을 가동시키는데 필요한 SW 모듈들로 이루어진 것은 다음의무엇인가? | 문서편집기 (Text edits)                | 프로그램 목록 관리기(Program library managers) | 결합편집과 실행기 (Linkage editors and loaders) | 버그 제거와 개발 보조(Debuggers and development aids) | NULL | @ 기계 명령어 응용 프로그램 버전을 실행하기 위해 필요한 소프트웨어 모듈을 어셈블하는유틸리티 프로그램은 링크 에디터와 로더이다.  | C    | 0     | 4    | 20040717  |           1 |
| 1998AA |   61 | D3-15 | 정보통신 시스템과 관련된 문장이 아닌 것은?                                       | 복수 층(multiple layers)으로 되어 있다| 운용 시스템(Operating system)으로 연계된다     | 정보의 내용에 따라 운용된다                   | 두 지점 사이를 정확히 전송하는 것과 관련한다          | NULL | @ 데이터 통신 시스템은 정보의 내용에 따라 운용되지 않는다                                                       | C    | 0     | 4    | 20040717  |           1 
|+--------+------+-------+----------------------------------------------------------------------------------------------+----------------------------------------+------------------------------------------------+-------------------------------------------------+-------------------------------------------------------+------+------------------------------------------------------------------------------------------------------------------------------------+------+-------+------+-----------+-------------+
2 rows in set (0.06 sec)

mysql> select * from vq01_log limit 2;
+------+----------+-------+--------------+-------------------------------+---------+------+------+------+------+-------------------------------------------------+------+-------+------+-----------+-------------+
| txt  | subj     | seq   | cat          | quiz                          | a   | b    | c    | d    | e    | hint | ans  | resid | mode | updateymd | updatechasu |
+------+----------+-------+--------------+-------------------------------+---------+------+------+------+------+-------------------------------------------------+------+-------+------+-----------+-------------+
| NULL | 한자     |    46 | 01_8급:0046  | 八                            | 여덟팔 | NULL | NULL | NULL | NULL | ㉠여덟 | 1    | 0     | 1    | 20050123  |           1 |
| NULL | 구약성경 | 20470 | C구:26겔 1:6 | 각각 네 얼굴과 네 날개가 있고 | NULL   | NULL | NULL | NULL | NULL | but each of them had four faces and four wings. | O    | 0     | 2    | 20050116  |           1 |
+------+----------+-------+--------------+-------------------------------+---------+------+------+------+------+-------------------------------------------------+------+-------+------+-----------+-------------+
2 rows in set (0.02 sec)

===========================성경추가예=================
insert into ts01 values('구약:01','구약성경01-창세기');
insert into ts01 values('구약:02','구약성경02-출에굽기');
insert into ts01 values('구약:03','구약성경03-레위기');
insert into ts01 values('구약:04','구약성경04-민수기');
insert into ts01 values('구약:05','구약성경05-신명기');
insert into ts01 values('구약:06','구약성경06-여호수아');
insert into ts01 values('구약:07','구약성경07-사사기');
insert into ts01 values('구약:08','구약성경08-롯기');
insert into ts01 values('구약:09','구약성경09-사무엘상');
insert into ts01 values('구약:10','구약성경10-사무엘하');
insert into ts01 values('구약:11','구약성경11-열왕기상');
insert into ts01 values('구약:12','구약성경12-열왕기하');
insert into ts01 values('구약:13','구약성경13-역대상');
insert into ts01 values('구약:14','구약성경14-역대하');
insert into ts01 values('구약:15','구약성경15-에스라');
insert into ts01 values('구약:16','구약성경16-느헤미아');
insert into ts01 values('구약:17','구약성경17-에스더');
insert into ts01 values('구약:18','구약성경18-욥기');
insert into ts01 values('구약:19','구약성경19-시편');
insert into ts01 values('구약:20','구약성경20-잠언');
insert into ts01 values('구약:21','구약성경21-전도서');
insert into ts01 values('구약:22','구약성경22-아가');
insert into ts01 values('구약:23','구약성경23-이사야');
insert into ts01 values('구약:24','구약성경24-예레미아');
insert into ts01 values('구약:25','구약성경25-예레미애가');
insert into ts01 values('구약:26','구약성경26-에스겔');
insert into ts01 values('구약:27','구약성경27-다니엘');
insert into ts01 values('구약:28','구약성경28-호세아');
insert into ts01 values('구약:29','구약성경29-요엘');
insert into ts01 values('구약:30','구약성경30-아모스');
insert into ts01 values('구약:31','구약성경31-오바댜');
insert into ts01 values('구약:32','구약성경32-요나');
insert into ts01 values('구약:33','구약성경33-미가');
insert into ts01 values('구약:34','구약성경34-나훔');
insert into ts01 values('구약:35','구약성경35-하박국');
insert into ts01 values('구약:36','구약성경36-스바냐');
insert into ts01 values('구약:37','구약성경37-학개');
insert into ts01 values('구약:38','구약성경38-스가락');
insert into ts01 values('구약:39','구약성경39-말라기');
insert into ts01 values('신약:40','신약성경40-마태복음');
insert into ts01 values('신약:41','신약성경41-마가복음');
insert into ts01 values('신약:42','신약성경42-누가복음');
insert into ts01 values('신약:43','신약성경43-요한복음');
insert into ts01 values('신약:44','신약성경44-사도행전');
insert into ts01 values('신약:45','신약성경45-로마서');
insert into ts01 values('신약:46','신약성경46-고린도전서');
insert into ts01 values('신약:47','신약성경47-고린도후서');
insert into ts01 values('신약:48','신약성경48-갈라디아서');
insert into ts01 values('신약:49','신약성경49-에베소서');
insert into ts01 values('신약:50','신약성경50-빌립보서');
insert into ts01 values('신약:51','신약성경51-골로세서');
insert into ts01 values('신약:52','신약성경52-데살로니가전서');
insert into ts01 values('신약:53','신약성경53-데살로니가후서');
insert into ts01 values('신약:54','신약성경54-디모데전서');
insert into ts01 values('신약:55','신약성경55-디모데후서');
insert into ts01 values('신약:56','신약성경56-디도서');
insert into ts01 values('신약:57','신약성경57-빌레몬서');
insert into ts01 values('신약:58','신약성경58-히브리서');
insert into ts01 values('신약:59','신약성경59-야고보서');
insert into ts01 values('신약:60','신약성경60-베드로전서');
insert into ts01 values('신약:61','신약성경61-베드로후서');
insert into ts01 values('신약:62','신약성경62-요한일서');
insert into ts01 values('신약:63','신약성경63-요한이서');
insert into ts01 values('신약:64','신약성경64-요한삼서');
insert into ts01 values('신약:65','신약성경65-유다서');
insert into ts01 values('신약:66','신약성경66-요한계시록');

insert into tp01 values('1301','구약:01','subj=''구약성경01-창세기''|(select * from tq07 where cat like ''_구:01%'')',0,0);
insert into tp01 values('1302','구약:02','subj=''구약성경02-출에굽기''|(select * from tq07 where cat like ''_구:02%'')',0,0);
insert into tp01 values('1303','구약:03','subj=''구약성경03-레위기''|(select * from tq07 where cat like ''_구:03%'')',0,0);
insert into tp01 values('1304','구약:04','subj=''구약성경04-민수기''|(select * from tq07 where cat like ''_구:04%'')',0,0);
insert into tp01 values('1305','구약:05','subj=''구약성경05-신명기''|(select * from tq07 where cat like ''_구:05%'')',0,0);
insert into tp01 values('1306','구약:06','subj=''구약성경06-여호수아''|(select * from tq07 where cat like ''_구:06%'')',0,0);
insert into tp01 values('1307','구약:07','subj=''구약성경07-사사기''|(select * from tq07 where cat like ''_구:07%'')',0,0);
insert into tp01 values('1308','구약:08','subj=''구약성경08-롯기''|(select * from tq07 where cat like ''_구:08%'')',0,0);
insert into tp01 values('1309','구약:09','subj=''구약성경09-사무엘상''|(select * from tq07 where cat like ''_구:09%'')',0,0);
insert into tp01 values('1310','구약:10','subj=''구약성경10-사무엘하''|(select * from tq07 where cat like ''_구:10%'')',0,0);
insert into tp01 values('1311','구약:11','subj=''구약성경11-열왕기상''|(select * from tq07 where cat like ''_구:11%'')',0,0);
insert into tp01 values('1312','구약:12','subj=''구약성경12-열왕기하''|(select * from tq07 where cat like ''_구:12%'')',0,0);
insert into tp01 values('1313','구약:13','subj=''구약성경13-역대상''|(select * from tq07 where cat like ''_구:13%'')',0,0);
insert into tp01 values('1314','구약:14','subj=''구약성경14-역대하''|(select * from tq07 where cat like ''_구:14%'')',0,0);
insert into tp01 values('1315','구약:15','subj=''구약성경15-에스라''|(select * from tq07 where cat like ''_구:15%'')',0,0);
insert into tp01 values('1316','구약:16','subj=''구약성경16-느헤미아''|(select * from tq07 where cat like ''_구:16%'')',0,0);
insert into tp01 values('1317','구약:17','subj=''구약성경17-에스더''|(select * from tq07 where cat like ''_구:17%'')',0,0);
insert into tp01 values('1318','구약:18','subj=''구약성경18-욥기''|(select * from tq07 where cat like ''_구:18%'')',0,0);
insert into tp01 values('1319','구약:19','subj=''구약성경19-시편''|(select * from tq07 where cat like ''_구:19%'')',0,0);
insert into tp01 values('1320','구약:20','subj=''구약성경20-잠언''|(select * from tq07 where cat like ''_구:20%'')',0,0);
insert into tp01 values('1321','구약:21','subj=''구약성경21-전도서''|(select * from tq07 where cat like ''_구:21%'')',0,0);
insert into tp01 values('1322','구약:22','subj=''구약성경22-아가''|(select * from tq07 where cat like ''_구:22%'')',0,0);
insert into tp01 values('1323','구약:23','subj=''구약성경23-이사야''|(select * from tq07 where cat like ''_구:23%'')',0,0);
insert into tp01 values('1324','구약:24','subj=''구약성경24-예레미아''|(select * from tq07 where cat like ''_구:24%'')',0,0);
insert into tp01 values('1325','구약:25','subj=''구약성경25-예레미애가''|(select * from tq07 where cat like ''_구:25%'')',0,0);
insert into tp01 values('1326','구약:26','subj=''구약성경26-에스겔''|(select * from tq07 where cat like ''_구:26%'')',0,0);
insert into tp01 values('1327','구약:27','subj=''구약성경27-다니엘''|(select * from tq07 where cat like ''_구:27%'')',0,0);
insert into tp01 values('1328','구약:28','subj=''구약성경28-호세아''|(select * from tq07 where cat like ''_구:28%'')',0,0);
insert into tp01 values('1329','구약:29','subj=''구약성경29-요엘''|(select * from tq07 where cat like ''_구:29%'')',0,0);
insert into tp01 values('1330','구약:30','subj=''구약성경30-아모스''|(select * from tq07 where cat like ''_구:30%'')',0,0);
insert into tp01 values('1331','구약:31','subj=''구약성경31-오바댜''|(select * from tq07 where cat like ''_구:31%'')',0,0);
insert into tp01 values('1332','구약:32','subj=''구약성경32-요나''|(select * from tq07 where cat like ''_구:32%'')',0,0);
insert into tp01 values('1333','구약:33','subj=''구약성경33-미가''|(select * from tq07 where cat like ''_구:33%'')',0,0);
insert into tp01 values('1334','구약:34','subj=''구약성경34-나훔''|(select * from tq07 where cat like ''_구:34%'')',0,0);
insert into tp01 values('1335','구약:35','subj=''구약성경35-하박국''|(select * from tq07 where cat like ''_구:35%'')',0,0);
insert into tp01 values('1336','구약:36','subj=''구약성경36-스바냐''|(select * from tq07 where cat like ''_구:36%'')',0,0);
insert into tp01 values('1337','구약:37','subj=''구약성경37-학개''|(select * from tq07 where cat like ''_구:37%'')',0,0);
insert into tp01 values('1338','구약:38','subj=''구약성경38-스가락''|(select * from tq07 where cat like ''_구:38%'')',0,0);
insert into tp01 values('1339','구약:39','subj=''구약성경39-말라기''|(select * from tq07 where cat like ''_구:39%'')',0,0);
insert into tp01 values('1340','신약:40','subj=''신약성경40-마태복음''|(select * from tq07 where cat like ''_신:40%'')',0,0);
insert into tp01 values('1341','신약:41','subj=''신약성경41-마가복음''|(select * from tq07 where cat like ''_신:41%'')',0,0);
insert into tp01 values('1342','신약:42','subj=''신약성경42-누가복음''|(select * from tq07 where cat like ''_신:42%'')',0,0);
insert into tp01 values('1343','신약:43','subj=''신약성경43-요한복음''|(select * from tq07 where cat like ''_신:43%'')',0,0);
insert into tp01 values('1344','신약:44','subj=''신약성경44-사도행전''|(select * from tq07 where cat like ''_신:44%'')',0,0);
insert into tp01 values('1345','신약:45','subj=''신약성경45-로마서''|(select * from tq07 where cat like ''_신:45%'')',0,0);
insert into tp01 values('1346','신약:46','subj=''신약성경46-고린도전서''|(select * from tq07 where cat like ''_신:46%'')',0,0);
insert into tp01 values('1347','신약:47','subj=''신약성경47-고린도후서''|(select * from tq07 where cat like ''_신:47%'')',0,0);
insert into tp01 values('1348','신약:48','subj=''신약성경48-갈라디아서''|(select * from tq07 where cat like ''_신:48%'')',0,0);
insert into tp01 values('1349','신약:49','subj=''신약성경49-에베소서''|(select * from tq07 where cat like ''_신:49%'')',0,0);
insert into tp01 values('1350','신약:50','subj=''신약성경50-빌립보서''|(select * from tq07 where cat like ''_신:50%'')',0,0);
insert into tp01 values('1351','신약:51','subj=''신약성경51-골로세서''|(select * from tq07 where cat like ''_신:51%'')',0,0);
insert into tp01 values('1352','신약:52','subj=''신약성경52-데살로니가전서''|(select * from tq07 where cat like ''_신:52%'')',0,0);
insert into tp01 values('1353','신약:53','subj=''신약성경53-데살로니가후서''|(select * from tq07 where cat like ''_신:53%'')',0,0);
insert into tp01 values('1354','신약:54','subj=''신약성경54-디모데전서''|(select * from tq07 where cat like ''_신:54%'')',0,0);
insert into tp01 values('1355','신약:55','subj=''신약성경55-디모데후서''|(select * from tq07 where cat like ''_신:55%'')',0,0);
insert into tp01 values('1356','신약:56','subj=''신약성경56-디도서''|(select * from tq07 where cat like ''_신:56%'')',0,0);
insert into tp01 values('1357','신약:57','subj=''신약성경57-빌레몬서''|(select * from tq07 where cat like ''_신:57%'')',0,0);
insert into tp01 values('1358','신약:58','subj=''신약성경58-히브리서''|(select * from tq07 where cat like ''_신:58%'')',0,0);
insert into tp01 values('1359','신약:59','subj=''신약성경59-야고보서''|(select * from tq07 where cat like ''_신:59%'')',0,0);
insert into tp01 values('1360','신약:60','subj=''신약성경60-베드로전서''|(select * from tq07 where cat like ''_신:60%'')',0,0);
insert into tp01 values('1361','신약:61','subj=''신약성경61-베드로후서''|(select * from tq07 where cat like ''_신:61%'')',0,0);
insert into tp01 values('1362','신약:62','subj=''신약성경62-요한일서''|(select * from tq07 where cat like ''_신:62%'')',0,0);
insert into tp01 values('1363','신약:63','subj=''신약성경63-요한이서''|(select * from tq07 where cat like ''_신:63%'')',0,0);
insert into tp01 values('1364','신약:64','subj=''신약성경64-요한삼서''|(select * from tq07 where cat like ''_신:64%'')',0,0);
insert into tp01 values('1365','신약:65','subj=''신약성경65-유다서''|(select * from tq07 where cat like ''_신:65%'')',0,0);
insert into tp01 values('1366','신약:66','subj=''신약성경66-요한계시록''|(select * from tq07 where cat like ''_신:66%'')',0,0);

insert into ts02 values('구약:01'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:02'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:03'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:04'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:05'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:06'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:07'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:08'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:09'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:10'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:11'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:12'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:13'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:14'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:15'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:16'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:17'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:18'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:19'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:20'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:21'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:22'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:23'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:24'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:25'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:26'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:27'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:28'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:29'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:30'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:31'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:32'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:33'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:34'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:35'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:36'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:37'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:38'        ,'박규선','20070404','21000405');
insert into ts02 values('구약:39'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:40'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:41'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:42'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:43'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:44'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:45'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:46'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:47'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:48'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:49'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:50'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:51'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:52'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:53'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:54'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:55'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:56'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:57'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:58'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:59'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:60'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:61'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:62'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:63'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:64'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:65'        ,'박규선','20070404','21000405');
insert into ts02 values('신약:66'        ,'박규선','20070404','21000405');

================================================================================