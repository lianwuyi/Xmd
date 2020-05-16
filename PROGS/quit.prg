
*!*	IF messagebox('您确定退出系统?',68,'删除') = 7
*!*	  WAIT CLEAR    
*!*	  RETURN
*!*	ENDIF

CLOSE DATABASES ALL 
CLOSE TABLES ALL 

QUIT 
