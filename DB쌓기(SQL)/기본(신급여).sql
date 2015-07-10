# 지급데이터베이스 글자깨짐 없애기
create database jigup DEFAULT CHARACTER SET utf8 collate utf8_general_ci;

# 지급데이터베이스로 이동
use jigup;

# DB_jigup 테이블 만들기
	create table jigup.db_jigup_new
	   (Id varchar(10) ,
		category varchar(50) ,
		Num_dec varchar(50) ,
		rrno_pat varchar(50) ,
		rrno_pat_seq varchar(50) ,
		name_pat varchar(50) ,
		rrno_ins varchar(50) ,
		rrno_ins_seq varchar(50) ,
		name_ins varchar(50) ,
		jung varchar(50) ,
		jupsu varchar(50) ,
		sort_jupsu varchar(50) ,
		Astatus varchar(50) ,
		amt varchar(50) , 
		rrno_Acc varchar(50) ,
		AccOwner varchar(50) ,
		bank varchar(50) ,
		account varchar(50) ,
		AccOwner_real varchar(50) ,
		pay_date varchar(50) ,
		rel varchar(50) ,
		stranger varchar(50),
		confirm varchar(50))
	DEFAULT CHARACTER SET utf8 collate utf8_general_ci;

	LOAD DATA LOCAL INFILE  'D:/DB/all.csv'
	INTO TABLE jigup.db_jigup_new
	FIELDS TERMINATED BY ',' 
	ENCLOSED BY '"'
	LINES TERMINATED BY '\r\n';
	
	# 자료검증 
	
		# 자료상태 확인
			select distinct Id from jigup.db_jigup_new;
			select distinct category from jigup.db_jigup_new;
			select distinct Num_dec from jigup.db_jigup_new;
			select distinct rrno_pat from jigup.db_jigup_new;
			select distinct rrno_pat_seq from jigup.db_jigup_new;
			select distinct name_pat from jigup.db_jigup_new;
			select distinct rrno_ins from jigup.db_jigup_new;
			select distinct rrno_ins_seq from jigup.db_jigup_new;
			select distinct name_ins from jigup.db_jigup_new;
			select distinct jung from jigup.db_jigup_new;
			select distinct jupsu from jigup.db_jigup_new;
			select distinct sort_jupsu from jigup.db_jigup_new;
			select distinct Astatus from jigup.db_jigup_new;
			select distinct amt from jigup.db_jigup_new;
			select distinct rrno_Acc from jigup.db_jigup_new;
			select distinct AccOwner from jigup.db_jigup_new;
			select distinct bank from jigup.db_jigup_new;
			select distinct account from jigup.db_jigup_new order by account;
			select distinct AccOwner_real from jigup.db_jigup_new order by AccOwner_real;
			select distinct pay_date from jigup.db_jigup_new order by pay_date;
			select distinct rel from jigup.db_jigup_new order by rel;
			select distinct stranger from jigup.db_jigup_new;
			select distinct confirm from jigup.db_jigup_new;
			
		# 이상자료 삭제
			# 지급결과 이상자료 삭제
				delete from jigup.db_jigup_new where category not in("반송수정","최초접수");
			# 공백 데이터 삭제
				delete from jigup.db_jigup_new where jigup.db_jigup_new.Astatus not in("실시간지급완료","계좌송금");
			# 지사연계계좌 삭제
				delete from jigup.db_jigup_new where left(rrno_Acc,10)="999999-999";
			# 예금주 공백 데이터 삭제
				delete from jigup.db_jigup_new where AccOwner="";
			# 은행명 공백 데이터 삭제		
				delete from jigup.db_jigup_new where bank="";
			
# 자료추가
	LOAD DATA LOCAL INFILE  'D:/db/jigup_add(201506).csv'
	INTO TABLE jigup.db_jigup_new
	FIELDS TERMINATED BY ',' 
	ENCLOSED BY '"'
	LINES TERMINATED BY '\r\n';

	# 공백 데이터 삭제
		delete from jigup.db_jigup_new where jigup.db_jigup_new.Astatus not in("실시간완료","계좌송금");
	# 지사연계계좌 삭제
		delete from jigup.db_jigup_new where left(rrno_Acc,10)="999999-999";
		select * from jigup.db_jigup_new where left(AccOwner,2)="건보";
	# 예금주 공백 데이터 삭제
		delete from jigup.db_jigup_new where AccOwner="";
	# 은행명 공백 데이터 삭제		
		delete from jigup.db_jigup_new where bank="";
		
# 계좌연계용 데이터
	# 접수방법 종류 뽑아내기
		select distinct sort_jupsu from jigup.db_jigup_new;

# db_jigup_new1 만들기
	create table jigup.db_jigup_new1 like jigup.db_jigup_new;
	insert into jigup.db_jigup_new1 select distinct * from jigup.db_jigup_new;

	# 인덱스 만들기
		alter table jigup.db_jigup_new1 add column bank_id text(3);
		alter table jigup.db_jigup_new1 add index (rrno_acc);
		alter table jigup.db_jigup_new1 add index (rrno_pat);
		alter table jigup.db_jigup_new1 add index (rrno_ins);

	# 은행코드 넣기
		update jigup.db_jigup_new1 set bank_id =case
			when bank='NH농협은행' then '011'
			when bank='국민은행' then '004'
			when bank='지역농.축협' then '012'
			when bank='우리은행' then '020'
			when bank='신한은행' then '088'
			when bank='기업은행' then '003'
			when bank='하나은행' then '081'
			when bank='외환은행' then '005'
			when bank='우체국' then '071'
			when bank='대구은행' then '031'
			when bank='부산은행' then '032'
			when bank='새마을금고중앙회' then '045'
			when bank='스탠다드차타드은행' then '023'
			when bank='광주은행' then '034'
			when bank='경남은행' then '039'
			when bank='씨티은행' then '027'
			when bank='전북은행' then '037'
			when bank='신협중앙회' then '048'
			when bank='수협중앙회' then '007'
			when bank='새마을금고연합회' then '045'
			when bank='새마을금고중앙회' then '045'
			when bank='제주은행' then '035'
			when bank='산업은행' then '002'
			when bank='조흥은행' then '088'
			when bank='유안타증권(구 동양)' then '209'
			when bank='상호저축은행' then '050'
			when bank='삼성증권' then '240'
			when bank='하이투자증권' then '262'
			when bank='현대증권' then '218'
			when bank='한화증권' then '269'
			when bank='산림조합' then '064'
			when bank='동양종합금융증권' then '209'
			when bank='메리츠증권' then '287'
			when bank='신한금융투자' then '278'
			when bank='미래에셋증권' then '230'
			when bank='에이치엠씨투자증권' then '263'
			when bank='한국투자증권' then '243'
			when bank='대우증권' then '238'
			when bank='하나대투증권' then '270'
			when bank='대신증권' then '267'
			when bank='SK증권' then '266'
			when bank='우리투자증권' then '247'
			when bank='홍콩은행' then '054'
			when bank='유진투자증권' then '280'
			when bank='신영증권' then '291'
			when bank='동부증권' then '279'
			when bank='NH투자증권' then '289'
			when bank='부국증권' then '290'
		end;

	#은행코드,접수구분 검증관련
		select * from jigup.db_jigup_new1 where bank_id is null;
		select distinct bank, bank_id from jigup.db_jigup_new1 limit 100;
		select distinct sort_jupsu from jigup.db_jigup_new1 limit 100;
		select * from jigup.db_jigup_new limit 10000;

		create table jigup.db_jigup_jisa like jigup.db_jigup_new1;
		insert into jigup.db_jigup_jisa select * from jigup.db_jigup_new1 where substring(Num_Dec,17,4)="0147";

# 미접수건 관련 DB만들기
	create table jigup.db_unpaid
	   (Num varchar(50) ,
		Num_Dec varchar(50) ,
		amt_first varchar(50) ,
		amt_last varchar(50) ,
		name_pat varchar(50) ,
		rrno_pat varchar(50) ,
		rrno_pat_seq varchar(50),
		jung_Dec varchar(50),
		Ent_Int varchar(50),
		Ent_Dec varchar(50),
		Date_rep_first varchar(50),
		Date_Due varchar(50),
		Reason_Change varchar(50),
		Astatus varchar(50),
		rrno_ins varchar(50),
		rrno_ins_seq varchar(50),
		name_ins varchar(50),
		date_acq_first varchar(50),
		date_acq_last varchar(50),
		tel_ins varchar(50),
		jisa_last1 varchar(50),
		jung_last varchar(50),
		Ent_tong varchar(50),
		Ent_Dec1 varchar(50),
		Sort_Ins varchar(50),
		Sort_Id varchar(50),
		Status_ins varchar(50),	
		Cell_Ins varchar(50),
		Jisa_last varchar(50),
        Ent_last_jagyuk varchar(50),
        Tel_Ent varchar(50),
		Fax_Ent varchar(50),
        Date_Ent_Enr varchar(50),
		Date_Ent_Ext varchar(50),
		EDI varchar(50),
		App_Ent varchar(50))
	DEFAULT CHARACTER SET utf8 collate utf8_general_ci;

	alter table jigup.db_unpaid add index (rrno_pat);
	alter table jigup.db_unpaid add index (rrno_ins);

	LOAD DATA LOCAL INFILE  'D:/DB/unpaid.csv'
		INTO TABLE jigup.db_unpaid
		FIELDS TERMINATED BY ',' 
		ENCLOSED BY '"'
		LINES TERMINATED BY '\r\n';
	


# 사망자 테이블 만들기
	Create table Num_Dec_Dead
		(
		Num_Dec1 varchar(50)
		)
		DEFAULT CHARACTER SET utf8 collate utf8_general_ci;

	LOAD DATA LOCAL INFILE 'D:/DB/dead.csv'
		INTO TABLE jigup.Num_Dec_Dead
		FIELDS TERMINATED BY ',' 
		ENCLOSED BY '"'
		LINES TERMINATED BY '\r\n';

	Create table rrno_Dead
		(
		rrno_pat varchar(50)
		)
	DEFAULT CHARACTER SET utf8 collate utf8_general_ci;

	insert into rrno_Dead
		(select rrno_pat from jigup.db_unpaid where Num_Dec in (select concat(left(trim(Num_Dec1),6),"-51-",right(trim(Num_Dec1),11)) from Num_Dec_Dead));



	# 연계용 표본테이블 만들기
	Create table link
		(
		Num_Dec1 char(19),
		jung char(11),
		Ent char(8),
		rrno_pat1 char(13),
		rrno_pat_seq1 char(2),
		name_pat char(20),
		rrno_ins1 char(13),
		rrno_ins_seq1 char(2),
		name_ins char(20),
		amt char(8),
		bank_id char(3),
		account char(20),
		Accowner char(20),
		rrno_acc1 char(13)
		)
		DEFAULT CHARACTER SET utf8 collate utf8_general_ci;



# 연계용 테이블 만들기
	create table link_pat_acc like link;
	create table link_ins_acc like link;
	create table link_jung_jung like link;
	create table link_pat_pat like link;	
	create table link_ins_ins like link;
	
# 연계용 테이블 비우기
	truncate table link_pat_acc;
	truncate table link_ins_acc;
	truncate table link_jung_jung;
	truncate table link_pat_pat;	
	truncate table link_ins_ins;
	
	
# 연계관련 테이블 연산

	#  수진자-예금주 텍스트접수
	insert into jigup.link_pat_acc(
	SELECT distinct concat(right(jigup.db_unpaid.Num_Dec,1),substring(jigup.db_unpaid.Num_Dec,17,4),left(jigup.db_unpaid.Num_Dec,6),51,substring(jigup.db_unpaid.Num_Dec,11,6)) as Num_Dec1,
		jigup.db_unpaid.jung_last as jung, jigup.db_unpaid.Ent_Dec1 as Ent, Replace(db_jigup_new1.rrno_pat,"-","") AS rrno_pat1, 
		"00" AS rrno_pat_seq1, jigup.db_unpaid.name_pat, Replace(db_jigup_new1.rrno_ins,"-","") AS rrno_ins1, 
		"00" AS rrno_ins_seq1, jigup.db_unpaid.name_ins, jigup.db_unpaid.amt_last as amt, db_jigup_new1.bank_id, db_jigup_new1.account, 
		db_jigup_new1.AccOwner, Replace(db_jigup_new1.rrno_ins,"-","") AS rrno_Acc1
		FROM jigup.db_unpaid INNER JOIN db_jigup_new1 ON jigup.db_unpaid.rrno_pat = db_jigup_new1.rrno_Acc
		WHERE (((db_jigup_new1.sort_jupsu) In ("유선접수","서면접수","팩스접수(2015년 이전)","인터넷접수","방문접수","고객센터","고객센터이관","사업장팩스","가입자팩스","경리단","현금지급","보상상한","자동이체","E.D.I접수","보험료환급금","공상","기타징수계좌","장기요양환급금계좌접수","현금급여환급금계좌접수"))) 
		ORDER BY db_jigup_new1.pay_date DESC);

	# 가입자-예금주 텍스트접수
	insert into jigup.link_ins_acc(
	SELECT distinct concat(right(jigup.db_unpaid.Num_Dec,1),substring(jigup.db_unpaid.Num_Dec,17,4),left(jigup.db_unpaid.Num_Dec,6),51,substring(jigup.db_unpaid.Num_Dec,11,6)) as Num_Dec1,
		jigup.db_unpaid.jung_last as jung, jigup.db_unpaid.Ent_Dec1 as Ent, Replace(db_jigup_new1.rrno_pat,"-","") AS rrno_pat1, 
		"00" AS rrno_pat_seq1, jigup.db_unpaid.name_pat, Replace(db_jigup_new1.rrno_ins,"-","") AS rrno_ins1, 
		"00" AS rrno_ins_seq1, jigup.db_unpaid.name_ins, jigup.db_unpaid.amt_last as amt, db_jigup_new1.bank_id, db_jigup_new1.account, 
		db_jigup_new1.AccOwner, Replace(db_jigup_new1.rrno_ins,"-","") AS rrno_Acc1
		FROM jigup.db_unpaid INNER JOIN db_jigup_new1 ON jigup.db_unpaid.rrno_ins = db_jigup_new1.rrno_Acc
		WHERE (((db_jigup_new1.sort_jupsu) In ("유선접수","서면접수","팩스접수(2015년 이전)","인터넷접수","방문접수","고객센터","고객센터이관","사업장팩스","가입자팩스","경리단","현금지급","보상상한","자동이체","E.D.I접수","보험료환급금","공상","기타징수계좌","장기요양환급금계좌접수","현금급여환급금계좌접수"))) 
		ORDER BY db_jigup_new1.pay_date DESC);

	# 증번호 텍스트접수
	insert into jigup.link_jung_jung(
	SELECT  distinct concat(right(jigup.db_unpaid.Num_Dec,1),substring(jigup.db_unpaid.Num_Dec,17,4),left(jigup.db_unpaid.Num_Dec,6),51,substring(jigup.db_unpaid.Num_Dec,11,6)) as Num_Dec1, 
		jigup.db_unpaid.jung_last, jigup.db_unpaid.Ent_Dec1 as ent, Replace(jigup.db_jigup_new1.rrno_pat,"-","") AS rrno_pat1, 
		"00" AS rrno_pat_seq1, jigup.db_unpaid.name_pat, Replace(jigup.db_jigup_new1.rrno_ins,"-","") AS rrno_ins1, 
		"00" AS rrno_ins_seq1, jigup.db_unpaid.name_ins, jigup.db_unpaid.amt_last as amt, jigup.db_jigup_new1.bank_id, jigup.db_jigup_new1.account, 
		jigup.db_jigup_new1.AccOwner, Replace(jigup.db_jigup_new1.rrno_ins,"-","") AS rrno_Acc1
		FROM jigup.db_unpaid INNER JOIN jigup.db_jigup_new1 ON jigup.db_unpaid.jung_last = jigup.db_jigup_new1.jung
		WHERE (((jigup.db_jigup_new1.sort_jupsu) In ("유선접수","서면접수","팩스접수(2015년 이전)","인터넷접수","방문접수","고객센터","고객센터이관","사업장팩스","가입자팩스","경리단","현금지급","보상상한","자동이체","E.D.I접수","보험료환급금","공상","기타징수계좌","장기요양환급금계좌접수","현금급여환급금계좌접수")))
		ORDER BY jigup.db_jigup_new1.pay_date DESC);

	# 수진자-수진자 텍스트접수
	insert into jigup.link_pat_pat(
	SELECT  distinct concat(right(jigup.db_unpaid.Num_Dec,1),substring(jigup.db_unpaid.Num_Dec,17,4),left(jigup.db_unpaid.Num_Dec,6),51,substring(jigup.db_unpaid.Num_Dec,11,6)) as Num_Dec1, 
		jigup.db_unpaid.jung_last, jigup.db_unpaid.Ent_Dec1 as ent, Replace(db_jigup_new1.rrno_pat,"-","") AS rrno_pat1, 
		"00" AS rrno_pat_seq1, jigup.db_unpaid.name_pat, Replace(db_jigup_new1.rrno_ins,"-","") AS rrno_ins1, 
		"00" AS rrno_ins_seq1, jigup.db_unpaid.name_ins, jigup.db_unpaid.amt_last as amt, db_jigup_new1.bank_id, db_jigup_new1.account, 
		db_jigup_new1.AccOwner, Replace(db_jigup_new1.rrno_ins,"-","") AS rrno_Acc1
		FROM jigup.db_unpaid INNER JOIN db_jigup_new1 ON jigup.db_unpaid.rrno_pat = db_jigup_new1.rrno_pat
		WHERE (((db_jigup_new1.sort_jupsu) In ("유선접수","서면접수","팩스접수(2015년 이전)","인터넷접수","방문접수","고객센터","고객센터이관","사업장팩스","가입자팩스","경리단","현금지급","보상상한","자동이체","E.D.I접수","보험료환급금","공상","기타징수계좌","장기요양환급금계좌접수","현금급여환급금계좌접수")))
		ORDER BY db_jigup_new1.pay_date DESC);

	# 가입자-가입자 텍스트접수
	insert into jigup.link_ins_ins(
	SELECT  distinct concat(right(jigup.db_unpaid.Num_Dec,1),substring(jigup.db_unpaid.Num_Dec,17,4),left(jigup.db_unpaid.Num_Dec,6),51,substring(jigup.db_unpaid.Num_Dec,11,6)) as Num_Dec1, 
		jigup.db_unpaid.jung_last, jigup.db_unpaid.Ent_Dec1 as ent, Replace(db_jigup_new1.rrno_pat,"-","") AS rrno_pat1, 
		"00" AS rrno_pat_seq1, jigup.db_unpaid.name_pat, Replace(db_jigup_new1.rrno_ins,"-","") AS rrno_ins1, 
		"00" AS rrno_ins_seq1, jigup.db_unpaid.name_ins, jigup.db_unpaid.amt_last as amt, db_jigup_new1.bank_id, db_jigup_new1.account, 
		db_jigup_new1.AccOwner, Replace(db_jigup_new1.rrno_ins,"-","") AS rrno_Acc1
		FROM jigup.db_unpaid INNER JOIN db_jigup_new1 ON jigup.db_unpaid.rrno_ins = db_jigup_new1.rrno_ins
		WHERE (((db_jigup_new1.sort_jupsu) In ("유선접수","서면접수","팩스접수(2015년 이전)","인터넷접수","방문접수","고객센터","고객센터이관","사업장팩스","가입자팩스","경리단","현금지급","보상상한","자동이체","E.D.I접수","보험료환급금","공상","기타징수계좌","장기요양환급금계좌접수","현금급여환급금계좌접수")))
		ORDER BY db_jigup_new1.pay_date DESC);


# EDI 테이블 만들기
Create table jigup.EDI
(
		Num_Dec1 varchar(50),
		Sort_EDI varchar(50),
		jung varchar(50),
		Ent_Int varchar(50),
		Ent varchar(50),
		Ent_branch varchar(50),
		Name_pat varchar(50),
		rrno_pat varchar(50),
		rrno_pat_seq varchar(50),
		Name_ins varchar(50),
		rrno_ins varchar(50),
		rrno_ins_seq varchar(50),
		Name_Acc varchar(50),
		rrno_Acc varchar(50),
		Rel_Acc varchar(50),
		Date_cmp varchar(50),
		Date_Ord_cmp varchar(50),
		Amt varchar(50),
		Date_Due varchar(50),
		Date_rep varchar(50),
		bank_id varchar(50),
		bank_name varchar(50),
		Account varchar(50),
		Tel_Ent varchar(50),
		Astatus varchar(50),
		Result_Total varchar(50),
		Result_bank varchar(50),
		Result_Acc varchar(50),
		Result_rrnedio_acc varchar(50),
		result_name_acc varchar(50),
		Name_Acc_Real varchar(50)
	)
	DEFAULT CHARACTER SET utf8 collate utf8_general_ci;

LOAD DATA LOCAL INFILE  'D:/DB/EDI.csv'
	INTO TABLE jigup.EDI
	FIELDS TERMINATED BY ','
	LINES TERMINATED BY '\r\n';



# EDI 연계
create table jigup.Link_EDI like jigup.link;

insert into jigup.Link_EDI(
SELECT concat(right(jigup.db_unpaid.Num_Dec,1),substring(jigup.db_unpaid.Num_Dec,17,4),left(jigup.db_unpaid.Num_Dec,6),51,substring(jigup.db_unpaid.Num_Dec,11,6)) as Num_Dec1,
	jigup.db_unpaid.jung_last, jigup.db_unpaid.Ent_Dec1, Replace(EDI.rrno_pat,"-","") AS rrno_pat1, 
	"00" AS rrno_pat_seq1, jigup.db_unpaid.name_pat, Replace(EDI.rrno_ins,"-","") AS rrno_ins1, 
	"00" AS rrno_ins_seq1, jigup.db_unpaid.name_ins, jigup.db_unpaid.amt_last, EDI.bank_id, EDI.account, 
	EDI.Name_Acc, Replace(EDI.rrno_ins,"-","") AS rrno_Acc1
	FROM jigup.db_unpaid INNER JOIN EDI ON jigup.db_unpaid.rrno_pat = EDI.rrno_Acc
	ORDER BY EDI.Date_cmp DESC);

insert into jigup.Link_EDI(
SELECT concat(right(jigup.db_unpaid.Num_Dec,1),substring(jigup.db_unpaid.Num_Dec,17,4),left(jigup.db_unpaid.Num_Dec,6),51,substring(jigup.db_unpaid.Num_Dec,11,6)) as Num_Dec1,
	jigup.db_unpaid.jung_last, jigup.db_unpaid.Ent_Dec1, Replace(EDI.rrno_pat,"-","") AS rrno_pat1, 
	"00" AS rrno_pat_seq1, jigup.db_unpaid.name_pat, Replace(EDI.rrno_ins,"-","") AS rrno_ins1, 
	"00" AS rrno_ins_seq1, jigup.db_unpaid.name_ins, jigup.db_unpaid.amt_last, EDI.bank_id, EDI.account, 
	EDI.Name_Acc, Replace(EDI.rrno_ins,"-","") AS rrno_Acc1
	FROM jigup.db_unpaid INNER JOIN EDI ON jigup.db_unpaid.rrno_ins = EDI.rrno_Acc
	ORDER BY EDI.Date_cmp DESC);

	
	
	
# 연계용 수진자_예금주 연계관련 임시테이블 만들기
	truncate temp;
	set profiling=1;

	insert into temp(
		SELECT  db_unpaid.Num_Dec,db_unpaid.jung_last,db_unpaid.Ent_Dec1, 
				db_unpaid.rrno_pat, db_unpaid.rrno_pat_seq, db_unpaid.name_pat, 
				db_unpaid.rrno_ins, db_unpaid.rrno_ins_seq, db_unpaid.name_ins, 
				db_unpaid.amt_last, db_jigup_new1.bank_id, db_jigup_new1.account, 
				db_jigup_new1.AccOwner, db_jigup_new1.rrno_Acc, db_jigup_new1.sort_jupsu, db_jigup_new1.pay_date
				FROM db_unpaid left JOIN db_jigup_new1 ON db_unpaid.rrno_pat = db_jigup_new1.rrno_Acc);

		set profiling=0;
		explain extended (
		SELECT  db_unpaid.Num_Dec,db_unpaid.jung_last,db_unpaid.Ent_Dec1, 
				db_unpaid.rrno_pat, db_unpaid.rrno_pat_seq, db_unpaid.name_pat, 
				db_unpaid.rrno_ins, db_unpaid.rrno_ins_seq, db_unpaid.name_ins, 
				db_unpaid.amt_last, db_jigup_new1.bank_id, db_jigup_new1.account, 
				db_jigup_new1.AccOwner, db_jigup_new1.rrno_Acc, db_jigup_new1.sort_jupsu, db_jigup_new1.pay_date
				FROM db_unpaid left JOIN db_jigup_new1 ON db_unpaid.rrno_pat = db_jigup_new1.rrno_Acc);
		show profiles;

		set profiling=1;
		SELECT  db_unpaid.Num_Dec,db_unpaid.jung_last,db_unpaid.Ent_Dec1, 
				db_unpaid.rrno_pat, db_unpaid.rrno_pat_seq, db_unpaid.name_pat, 
				db_unpaid.rrno_ins, db_unpaid.rrno_ins_seq, db_unpaid.name_ins, 
				db_unpaid.amt_last, db_jigup_new1.bank_id, db_jigup_new1.account, 
				db_jigup_new1.AccOwner, db_jigup_new1.rrno_Acc, db_jigup_new1.sort_jupsu, db_jigup_new1.pay_date
				FROM db_unpaid left JOIN db_jigup_new1 ON db_unpaid.rrno_pat = db_jigup_new1.rrno_Acc;
		show profiles;

		select count(*) from db_jigup_new1 where (sort_jupsu in ("팩스접수(2015년 이전)","방문접수","유선접수","사업장팩스","가입자팩스")) group by left(num_dec,4);
	create table jigup.db_jigup_add
	   (Id varchar(10) ,
		category varchar(50) ,
		Num_dec varchar(50) ,
		rrno_pat varchar(50) ,
		rrno_pat_seq varchar(50) ,
		name_pat varchar(50) ,
		rrno_ins varchar(50) ,
		rrno_ins_seq varchar(50) ,
		name_ins varchar(50) ,
		jung varchar(50) ,
		jupsu varchar(50) ,
		sort_jupsu varchar(50) ,
		Astatus varchar(50) ,
		amt varchar(50) , 
		rrno_Acc varchar(50) ,
		AccOwner varchar(50) ,
		bank varchar(50) ,
		account varchar(50) ,
		AccOwner_real varchar(50) ,
		pay_date varchar(50) ,
		rel varchar(50) ,
		stranger varchar(50),
		confirm varchar(50))
	DEFAULT CHARACTER SET utf8 collate utf8_general_ci;
	
		# 연계대상 입력테이블 만들기
		create table jigup.link
		   (rrno_Acc varchar(50) ,
			AccOwner varchar(50) ,
			bank varchar(50) ,
			account varchar(50) ,
			sort_jupsu varchar(50),
			AccOwner_real varchar(50)
		)DEFAULT CHARACTER SET utf8 collate utf8_general_ci;
		insert into jigup.link SELECT distinct rrno_Acc, AccOwner,bank,account,sort_jupsu,AccOwner_real
		FROM jigup.db_jigup_new where (substring(Num_Dec,17,4)="0147") and (rrno_Acc<>rrno_pat) and (sort_jupsu in("고객센터","서면접수","팩스접수(2015년 이전)","E.D.I접수","인터넷접수","유선접수","방문접수"))
		order by Num_Dec desc;

	# 연계대상 데이터를 csv로 내보내기
		alter table jigup.link add column bank_id text(3);
		alter table jigup.link add index (rrno_Acc);
		update jigup.link set bank_id =case
			when bank='NH농협은행' then '011'
			when bank='국민은행' then '004'
			when bank='지역농.축협' then '012'
			when bank='우리은행' then '020'
			when bank='신한은행' then '088'
			when bank='기업은행' then '003'
			when bank='하나은행' then '081'
			when bank='외환은행' then '005'
			when bank='우체국' then '071'
			when bank='대구은행' then '031'
			when bank='부산은행' then '032'
			when bank='새마을금고중앙회' then '045'
			when bank='스탠다드차타드은행' then '023'
			when bank='광주은행' then '034'
			when bank='경남은행' then '039'
			when bank='씨티은행' then '027'
			when bank='전북은행' then '037'
			when bank='신협중앙회' then '048'
			when bank='수협중앙회' then '007'
			when bank='새마을금고연합회' then '045'
			when bank='새마을금고중앙회' then '045'
			when bank='제주은행' then '035'
			when bank='산업은행' then '002'
			when bank='조흥은행' then '088'
			when bank='유안타증권(구 동양)' then '209'
			when bank='상호저축은행' then '050'
			when bank='삼성증권' then '240'
			when bank='하이투자증권' then '262'
			when bank='현대증권' then '218'
			when bank='한화증권' then '269'
			when bank='산림조합' then '064'
			when bank='동양종합금융증권' then '209'
			when bank='메리츠증권' then '287'
			when bank='신한금융투자' then '278'
			when bank='미래에셋증권' then '230'
			when bank='에이치엠씨투자증권' then '263'
			when bank='한국투자증권' then '243'
			when bank='대우증권' then '238'
			when bank='하나대투증권' then '270'
			when bank='대신증권' then '267'
			when bank='SK증권' then '266'
			when bank='우리투자증권' then '247'
			when bank='홍콩은행' then '054'
			when bank='유진투자증권' then '280'
			when bank='신영증권' then '291'
			when bank='동부증권' then '279'
			when bank='NH투자증권' then '289'
			when bank='부국증권' then '290'
		end;
		SELECT distinct rrno_Acc, bank_id, account, AccOwner, sort_jupsu,AccOwner_real INTO OUTFILE "D:/db/link.csv"
			FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"'
			LINES TERMINATED BY "\n"
			FROM jigup.link;
		select distinct sort_jupsu from jigup.db_jigup_new1;