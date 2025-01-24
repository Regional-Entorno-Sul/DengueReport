Function main()

set century on
set date british

public sum_not2, sum_conf2, sum_obito_inve2, sum_obito2

screen()
cleanning()
screen()
verify()
screen()
extract()
screen()

@ 6,0 say "Transferindo os dados do arquivo dbf para um aquivo de texto..."
use "c:\DengueReport\dbf\dengon.dbf"
copy fields nu_notific, tp_not, sem_not, nu_ano, id_mn_resi, classi_fin, evolucao to "c:\DengueReport\run\dengue.txt" delimited

creator()
modelo()
municres()
notificado()
confirmado()
obito()
investiga()
csv2dbf( "obitos" )
csv2dbf( "notificado" )
csv2dbf( "confirmado" )
csv2dbf( "investigado" )
modelo2()

fill_modelo2( "obitos" )
fill_modelo2( "notificado" )
fill_modelo2( "confirmado" )
fill_modelo2( "investigado" )
fill_modelo2( "incidencia" )
total()
output()

function creator()
@ 7,0 say "Criando um database no Sqlite com os dados de Dengue..."
? ""
cCom := "c:\DengueReport\run\sqlite3.exe c:/DengueReport/run/dengue" + " " + chr(34) + ".read c:/DengueReport/run/creator.sql" + chr(34)
__run( cCom )
return

function modelo
@ 8,0 say "Criando um arquivo modelo no formato dbf para receber as queries...                              "
aStruct := { { "semana","C",6,0 }, ;
             { "count","C",12,0 }}
			 dbcreate ("c:\DengueReport\mod\modelo.dbf", aStruct)

use "c:\DengueReport\mod\modelo.dbf"

for n = 1 to 53

if n < 10
cString := HB_ArgV ( 1 ) + "0" + alltrim( str( n ) )
else
cString := HB_ArgV ( 1 ) + alltrim( str( n ) )
endif

append blank
replace semana with cString
next

close
return

function notificado()

@ 10,0 say "Criando as queries com os casos notificados de Dengue..."
nArraySize := len( aArray_mun_list )
for n := 1 to nArraySize 	

cCom2 := "c:\DengueReport\run\sqlite3.exe c:/DengueReport/run/dengue" + " " + chr(34) + ;
         ".param init" + chr(34) + " " + chr(34) + ".param set :munires ";
		 + chr(39) + alltrim( aArray_mun_list[n,1] ) + chr(39) + chr(34) + " " + chr(34) +;
		 ".read c:/DengueReport/run/notificado.sql" + chr(34)
__run( cCom2 )
cFile := "c:\DengueReport\run\notificados_" + alltrim( aArray_mun_list[n,1] ) + ".csv"
copy file "c:\DengueReport\run\notificados.csv" to ( cFile )

next

return

function municres()
@ 9,0 say "Transportando para um array os dados dos municipios em 'list_muns.txt'..."
copy file "c:\DengueReport\mod\list_muns.dbf" to "c:\DengueReport\set\list_muns.dbf"

use "c:\DengueReport\set\list_muns.dbf"
zap
append from "c:\DengueReport\set\list_muns.txt" delimited with '"'
close

	  public aArray_mun_list := {}
	  use "c:\DengueReport\set\list_muns.dbf"
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray_mun_list, {alltrim(cod_munres)} )
	  skip
	  NEXT
	  close

return

function confirmado()

@ 11,0 say "Criando as queries com os casos confirmados de Dengue..."
nArraySize := len( aArray_mun_list )
for n := 1 to nArraySize

cCom3 := "c:\DengueReport\run\sqlite3.exe c:/DengueReport/run/dengue" + " " + chr(34) + ;
         ".param init" + chr(34) + " " + chr(34) + ".param set :munires ";
		 + chr(39) + alltrim( aArray_mun_list[n,1] ) + chr(39) + chr(34) + " " + chr(34) +;
		 ".read c:/DengueReport/run/confirmado.sql" + chr(34)
__run( cCom3 )
cFile := "c:\DengueReport\run\confirmados_" + alltrim( aArray_mun_list[n,1] ) + ".csv"
copy file "c:\DengueReport\run\confirmados.csv" to ( cFile )

next

return

function obito()

@ 12,0 say "Criando as queries com os obitos de Dengue..."
nArraySize := len( aArray_mun_list )
for n := 1 to nArraySize

cCom4 := "c:\DengueReport\run\sqlite3.exe c:/DengueReport/run/dengue" + " " + chr(34) + ;
         ".param init" + chr(34) + " " + chr(34) + ".param set :munires ";
		 + chr(39) + alltrim( aArray_mun_list[n,1] ) + chr(39) + chr(34) + " " + chr(34) +;
		 ".read c:/DengueReport/run/obito.sql" + chr(34)
__run( cCom4 )
cFile := "c:\DengueReport\run\obitos_" + alltrim( aArray_mun_list[n,1] ) + ".csv"
copy file "c:\DengueReport\run\obitos.csv" to ( cFile )

next

return

function investiga()

@ 13,0 say "Criando as queries com os obitos de Dengue sob investigacao..."
nArraySize := len( aArray_mun_list )
for n := 1 to nArraySize

cCom5 := "c:\DengueReport\run\sqlite3.exe c:/DengueReport/run/dengue" + " " + chr(34) + ;
         ".param init" + chr(34) + " " + chr(34) + ".param set :munires ";
		 + chr(39) + alltrim( aArray_mun_list[n,1] ) + chr(39) + chr(34) + " " + chr(34) +;
		 ".read c:/DengueReport/run/investiga.sql" + chr(34)
__run( cCom5 )
cFile := "c:\DengueReport\run\investigados_" + alltrim( aArray_mun_list[n,1] ) + ".csv"
copy file "c:\DengueReport\run\investigados.csv" to ( cFile )

next

return

function csv2dbf( cTipo )

aStruct := { { "semana","C",6,0 }, ;
             { "count","C",12,0 }}
			 dbcreate ("c:\DengueReport\run\receive.dbf", aStruct)

if cTipo = "obitos"
@ 14,0 say "Transferindo os dados da query do formato CSV para DBF: obitos de Dengue..."
nArraySize := ADir( "c:\DengueReport\run\obitos_*.csv" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\obitos_*.csv" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

use "c:\DengueReport\run\receive.dbf"
append from ( cFile ) delimited

	  public aArray1 := {}
	  use "c:\DengueReport\run\receive.dbf"
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

copy file "c:\DengueReport\mod\modelo.dbf" to "c:\DengueReport\run\modelo.dbf"
use "c:\DengueReport\run\modelo.dbf"

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0 
replace count with cValor
else
replace count with "0"
endif
endif

skip
enddo

close

cFileName := aName[f]
cFileName2 := AtRepl( ".csv", cFileName, "" )
cFileOutput := "c:\DengueReport\run\modelo_" + cFileName2 + ".dbf"
copy file "c:\DengueReport\run\modelo.dbf" to ( cFileOutput )

use "c:\DengueReport\run\receive.dbf"
zap
close

delete file "c:\DengueReport\run\modelo.dbf"

next

endif

if cTipo = "notificado"
@ 15,0 say "Transferindo os dados da query do formato CSV para DBF: notificacoes de Dengue..."
nArraySize := ADir( "c:\DengueReport\run\notificados_*.csv" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\notificados_*.csv" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

use "c:\DengueReport\run\receive.dbf"
append from ( cFile ) delimited

	  public aArray1 := {}
	  use "c:\DengueReport\run\receive.dbf"
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

copy file "c:\DengueReport\mod\modelo.dbf" to "c:\DengueReport\run\modelo.dbf"
use "c:\DengueReport\run\modelo.dbf"

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0 
replace count with cValor
else
replace count with "0"
endif
endif

skip
enddo

close

cFileName := aName[f]
cFileName2 := AtRepl( ".csv", cFileName, "" )
cFileOutput := "c:\DengueReport\run\modelo_" + cFileName2 + ".dbf"
copy file "c:\DengueReport\run\modelo.dbf" to ( cFileOutput )

use "c:\DengueReport\run\receive.dbf"
zap
close

delete file "c:\DengueReport\run\modelo.dbf"

next

endif

if cTipo = "confirmado"
@ 16,0 say "Transferindo os dados da query do formato CSV para DBF: confirmacoes de Dengue..."
nArraySize := ADir( "c:\DengueReport\run\confirmados_*.csv" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\confirmados_*.csv" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

use "c:\DengueReport\run\receive.dbf"
append from ( cFile ) delimited

	  public aArray1 := {}
	  use "c:\DengueReport\run\receive.dbf"
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

copy file "c:\DengueReport\mod\modelo.dbf" to "c:\DengueReport\run\modelo.dbf"
use "c:\DengueReport\run\modelo.dbf"

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0 
replace count with cValor
else
replace count with "0"
endif
endif

skip
enddo

close

cFileName := aName[f]
cFileName2 := AtRepl( ".csv", cFileName, "" )
cFileOutput := "c:\DengueReport\run\modelo_" + cFileName2 + ".dbf"
copy file "c:\DengueReport\run\modelo.dbf" to ( cFileOutput )

use "c:\DengueReport\run\receive.dbf"
zap
close

delete file "c:\DengueReport\run\modelo.dbf"

next

endif

if cTipo = "investigado"
@ 17,0 say "Transferindo os dados da query do formato CSV para DBF: investigacao de obitos..."
nArraySize := ADir( "c:\DengueReport\run\investigados_*.csv" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\investigados_*.csv" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

use "c:\DengueReport\run\receive.dbf"
append from ( cFile ) delimited

	  public aArray1 := {}
	  use "c:\DengueReport\run\receive.dbf"
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

copy file "c:\DengueReport\mod\modelo.dbf" to "c:\DengueReport\run\modelo.dbf"
use "c:\DengueReport\run\modelo.dbf"

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0 
replace count with cValor
else
replace count with "0"
endif
endif

skip
enddo

close

cFileName := aName[f]
cFileName2 := AtRepl( ".csv", cFileName, "" )
cFileOutput := "c:\DengueReport\run\modelo_" + cFileName2 + ".dbf"
copy file "c:\DengueReport\run\modelo.dbf" to ( cFileOutput )

use "c:\DengueReport\run\receive.dbf"
zap
close

delete file "c:\DengueReport\run\modelo.dbf"

next

endif

return

function modelo2

@ 18,0 say "Criando arquivo modelo dbf para receber os resultados consolidados..."
aStruct := { { "semana","C",6,0 }, ;
             { "notificado","C",12,0 }, ;
             { "confirmado","C",12,0 }, ;
             { "incidencia","N",8,2 }, ;
             { "obito_inve","C",12,0 }, ;
             { "obito_conf","C",12,0 }}			 
			 dbcreate ("c:\DengueReport\mod\modelo2.dbf", aStruct)

use "c:\DengueReport\mod\modelo2.dbf"

for n = 1 to 53

if n < 10
cString := HB_ArgV ( 1 ) + "0" + alltrim( str( n ) )
else
cString := HB_ArgV ( 1 ) + alltrim( str( n ) )
endif

append blank
replace semana with cString
next

append blank
replace semana with "Total"

close

nArraySize := len( aArray_mun_list )
for n := 1 to nArraySize

cFile := "c:\DengueReport\run\modelo2_" + alltrim( aArray_mun_list[n,1] ) + ".dbf"
copy file "c:\DengueReport\mod\modelo2.dbf" to ( cFile )

next

return

function fill_modelo2( cTipo )

if cTipo = "obitos"
@ 19,0 say "Preenchendo o consolidado dos dados de obitos de Dengue..."

nArraySize := ADir( "c:\DengueReport\run\modelo2_*.dbf" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\modelo2_*.dbf" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

cSubStr := SubStr( cFile, 29, 6)

cFile2 := "c:\DengueReport\run\modelo_obitos_" + cSubStr + ".dbf"

	  public aArray1 := {}
	  use ( cFile2 )
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

use ( cFile )

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0
replace obito_conf with cValor
else
replace obito_conf with "0"
endif

endif

skip
enddo

close

next

endif

if cTipo = "notificado"
@ 20,0 say "Preenchendo o consolidado dos dados de notificacoes de Dengue..."

nArraySize := ADir( "c:\DengueReport\run\modelo2_*.dbf" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\modelo2_*.dbf" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

cSubStr := SubStr( cFile, 29, 6)

cFile2 := "c:\DengueReport\run\modelo_notificados_" + cSubStr + ".dbf"

	  public aArray1 := {}
	  use ( cFile2 )
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

use ( cFile )

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0
replace notificado with cValor
else
replace notificado with "0"
endif

endif

skip
enddo

close

next

endif

if cTipo = "confirmado"
@ 21,0 say "Preenchendo o consolidado dos dados de confirmacoes de Dengue..."

nArraySize := ADir( "c:\DengueReport\run\modelo2_*.dbf" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\modelo2_*.dbf" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

cSubStr := SubStr( cFile, 29, 6)

cFile2 := "c:\DengueReport\run\modelo_confirmados_" + cSubStr + ".dbf"

	  public aArray1 := {}
	  use ( cFile2 )
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

use ( cFile )

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0
replace confirmado with cValor
else
replace confirmado with "0"
endif

endif

skip
enddo

close

next

endif

if cTipo = "investigado"
@ 22,0 say "Preenchendo o consolidado dos dados de obitos sob investigacao..."

nArraySize := ADir( "c:\DengueReport\run\modelo2_*.dbf" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\modelo2_*.dbf" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

cSubStr := SubStr( cFile, 29, 6)

cFile2 := "c:\DengueReport\run\modelo_investigados_" + cSubStr + ".dbf"

	  public aArray1 := {}
	  use ( cFile2 )
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

use ( cFile )

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0
replace obito_inve with cValor
else
replace obito_inve with "0"
endif

endif

skip
enddo

close

next

endif

if cTipo = "incidencia"
@ 23,0 say "Preenchendo o consolidado dos dados de incidencia de Dengue..."

public nPop, cConf
nArraySize := ADir( "c:\DengueReport\run\modelo2_*.dbf" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\modelo2_*.dbf" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]

cSubStr := SubStr( cFile, 29, 6)

use "c:\DengueReport\tab\pop2021.dbf"
locate for left( full_code,6 ) = cSubStr
if found() = .T.
nPop := n_pop_
else
endif
close

cFile2 := "c:\DengueReport\run\modelo_confirmados_" + cSubStr + ".dbf"

	  public aArray1 := {}
	  use ( cFile2 )
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray1, {alltrim( semana ),alltrim( count )} )
	  skip
	  NEXT
	  close
	  aArray1 := asort( aArray1,,,{|x,y| x[1] < y[1]} )

use ( cFile )

nVezes := 0
do while .not. eof()
erro = 0
cValor = alltrim( semana )

if empty(cValor) = .F.
nPlace := ascan( aArray1, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray1[nPlace,2]
recover
erro = 1
end sequence

if erro = 0
replace incidencia with ( (val( cValor ) * 100000 ) / nPop )
else
replace incidencia with 0
endif

endif

skip
enddo

close

next

endif

return

function cleanning()

@ 6,0 say ""
__run( "del /F /Q c:\DengueReport\run\*.dbf" )
__run( "del /F /Q c:\DengueReport\run\*.csv" )
__run( "del /F /Q c:\DengueReport\dbf\*.dbf" )
__run( "del /F /Q c:\DengueReport\out\*.csv" )
delete file "c:\DengueReport\dbf\dengon.dbf"
delete file "c:\DengueReport\run\dengue"
delete file "c:\DengueReport\run\dengue.txt"
delete file "c:\DengueReport\run\modelo.dbf"
__run( "ren sed* sad*.123" )
__run( "ren sad.123 sed.exe" )
__run( "del /F /Q c:\DengueReport\exe\*.123" )

return

function verify()

nTotal := ADir ( "c:\DengueReport\dbf\*.zip" )

if nTotal <> 1
@ 6,0 say "Erro!" color "r+"
@ 7,0 say "Somente podera haver um arquivo no formato zip no diretorio 'DBF'."
@ 8,0 say "Fim do programa."
quit
else
@ 6,0 say "Quantidade de arquivos zip...ok"
endif

return

function screen()
set color to g+/
clear screen
@ 1,0 say "--------------------------------------------------------------------------" color "w+/b+"
@ 2,0 say " DengueReport.exe - versao 1.2 - 21/06/2024                               " color "w+/b+"
@ 3,0 say " Sintaxe: DengueReport.exe [ano]                                          " color "w+/b+"
@ 4,0 say " Exemplo: DengueReport.exe 2024                                           " color "w+/b+"
@ 5,0 say "--------------------------------------------------------------------------" color "w+/b+"
return

function extract()

@ 6,0 say "Extraindo o arquivo dbf do arquivo zip..."
nArraySize := ADir( "c:\DengueReport\dbf\" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\dbf\" ),aName )
for f = 1 to nArraySize
hb_UnzipFile( "c:\DengueReport\dbf\" + aName[f] )
next
@ 7,0 say "Arquivo zip:" color "w+/"
@ 7,13 say ( "c:\DengueReport\dbf\" + aName[1] ) color "gr+/"

nArraySize2 := ADir( "c:\DengueReport\dbf\*.dbf" )
aName2 := Array( nArraySize2 )
ADir( ( "c:\DengueReport\dbf\*.dbf" ),aName2 )

@ 8,0 say "Arquivo dbf:" color "w+/"
@ 8,13 say ( "c:\DengueReport\dbf\" + aName2[1] ) color "gr+/"

rename ( "c:\DengueReport\dbf\" + aName2[1] ) to ("c:\DengueReport\dbf\dengon.dbf")

return

function total()

@ 24,0 say "Calculando os totais..."
nArraySize0 := ADir( "c:\DengueReport\run\modelo2_*.dbf" )
aName0 := Array( nArraySize0 )
ADir( ( "c:\DengueReport\run\modelo2_*.dbf" ),aName0 )

for f = 1 to nArraySize0
cFile := "c:\DengueReport\run\"+aName0[f]

use ( cFile )

sum val( notificado ) to sum_not while .not. eof()
goto top
sum val( confirmado ) to sum_conf while .not. eof()
goto top
sum val( obito_inve ) to sum_obito_inve while .not. eof()
goto top
sum val( obito_conf ) to sum_obito while .not. eof()

sum_not2 := sum_not
sum_conf2 := sum_conf
sum_obito_inve2 := sum_obito_inve
sum_obito2 := sum_obito

close

use ( cFile )

goto 54
replace notificado with alltrim( str( sum_not2 ) )
replace confirmado with alltrim( str( sum_conf2 ) )
replace obito_inve with alltrim( str( sum_obito_inve2 ) )
replace obito_conf with alltrim( str( sum_obito2 ) )

close

next

return

function output()

@ 25,0 say "Convertendo os arquivos processados para o formato csv e transferindo para o diretorio 'out'..."

	  public aArray_list := {}
	  use ( "c:\DengueReport\set\list_muns.dbf" )
      nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray_list, {alltrim( cod_munres ),alltrim( munic_name )} )
	  skip
	  NEXT
	  close

nArraySize := ADir( "c:\DengueReport\run\modelo2_*.dbf" )
aName := Array( nArraySize )
ADir( ( "c:\DengueReport\run\modelo2_*.dbf" ),aName )

for f = 1 to nArraySize
cFile := "c:\DengueReport\run\"+aName[f]
cFile2 := AtRepl( "run", cFile, "out" )
cFile3 := AtRepl( "dbf", cFile2, "csv" )

use ( cFile )
copy to ( cFile3 ) delimited

cCommand := "c:\DengueReport\exe\sed.exe -i " + chr(34) + "1i semana,notificado,confirmado,incidencia,obito em investigacao,obito confirmado" + ;
chr(34) + " " + ( cFile3 )
__run( cCommand )

next

@ 26,0 say "Renomeando os arquivos transferidos com o nome do municipio..."

nArraySize2 := ADir( "c:\DengueReport\out\modelo2_*.csv" )
aName2 := Array( nArraySize2 )
ADir( ( "c:\DengueReport\out\modelo2_*.csv" ),aName2 )

for n = 1 to nArraySize2
cFile4 := "c:\DengueReport\out\"+aName2[n]
cString := substr( cFile4, 29, 6)

erro = 0
cValor = alltrim( cString )

if empty( cValor ) = .F.
nPlace := ascan( aArray_list, {|x| x[1] == (cValor) } )

begin sequence WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
cValor := aArray_list[nPlace,1]
recover
erro = 1
end sequence

if erro = 0
cFile5 := AtRepl( alltrim( cString ), cFile4, aArray_list[nPlace,2] )
rename ( cFile4 ) to ( cFile5 )
else
endif

endif

next

@ 27,0 say "Fim do processamento."
@ 28,0 say ""

return

return nil