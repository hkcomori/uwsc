Set-Variable uri 'http://hogehoge.com' -Option Constant
$userName = "userName"
$userPwd = "Password"

#トラップ
try{
	$ie = AutoIE::New( "$uri" )

	#DOMのID名などは適宜変更してください。
	$ie.SetValueById( "user_login", $userName )
	$ie.SetValueById( "user_pass", $userPwd )

	$ie.ClickById( "submit" )
}catch [Exception]{
    $Error
}finally{
    exit
}

#以降ログイン後の処理など


################################################################
# IEを自動操作するためのインタフェースを提供するクラス         #
################################################################
Class AutoIE
{
	[PSObject] $ie					#IEオブジェクト
	[string] $lastId = ""			#最後に操作した要素のid
	[string] $lastName = ""			#最後に操作した要素のname

	################################################################
	# コンストラクタ                                               #
	#  戻り値                                                      #
	#   なし                                                       #
	#  引数                                                        #
	#   $a_uri		開くURL                                        #
	################################################################
	IE( [string] $a_uri )
	{
		$this.ie = new-object -com InternetExplorer.Application
		$this.ie.visible = true
		$this.ie.navigate( $a_uri )
		While($ie.Busy){ start-sleep -milliseconds 100 }
	}

	################################################################
	# 指定IDの要素に値を設定する                                   #
	#  戻り値                                                      #
	#   なし                                                       #
	#  引数                                                        #
	#   $a_id		指定ID                                         #
	#   $a_val		設定値                                         #
	################################################################
	[void] SetValueByID( [string] $a_id [=$this.lastID], [string] $a_val )
	{
		$this.lastId = $a_id
		$_targetId = $this.ie.document.getElementById( $a_id )
		$_targetId.value = $a_val
		While($ie.Busy){ start-sleep -milliseconds 100 }
	}

	################################################################
	# 指定IDの要素に設定されている値を取得する                     #
	#  戻り値                                                      #
	#   [string]	設定値                                         #
	#  引数                                                        #
	#   $a_id		指定ID                                         #
	################################################################
	[string] GetValueByID( [string] $a_id [=$this.lastID] )
	{
		$this.lastId = $a_id
		$_targetId = $this.ie.document.getElementById( $a_id )
		return $_targetId.value
	}

	################################################################
	# 指定name、指定値のチェックボックスをチェックする             #
	#  戻り値                                                      #
	#   なし                                                       #
	#  引数                                                        #
	#   $a_id		指定name                                       #
	#   $a_val		指定値                                         #
	################################################################
	[void] CheckByName( [string] $a_name [=$this.lastName], [string] $a_val )
	{
		$this.lastName = $a_name
		$_targetName = $this.ie.document.getElementByName( $a_name )
		foreach( mshtml.HTMLInputElement ele in $_targetName ){
			if( $_targetName.value == $a_val ){
				$_targetName.setAttribute("checked", "checked")
			}
		}
		While($ie.Busy){ start-sleep -milliseconds 100 }
	}

	################################################################
	# 指定IDの要素をクリックする                                   #
	#  戻り値                                                      #
	#   なし                                                       #
	#  引数                                                        #
	#   $a_id		指定ID                                         #
	################################################################
	[void] ClickByID( [string] $a_id [=$this.lastID] )
	{
		$this.lastId = $a_id
		$_targetId = $this.ie.document.getElementById( $a_id )
		$_targetId.click()
		While($ie.Busy){ start-sleep -milliseconds 100 }
	}

	################################################################
	# IEの自動操作を終了する                                       #
	#  戻り値                                                      #
	#   なし                                                       #
	#  引数                                                        #
	#   なし                                                       #
	################################################################
	[void] Quit()
	{
		$this.quit()
	}
}
