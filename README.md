#WOT地图对照生成
#WOT-Map-Information-Creater  

##一个从坦克世界客户端文件中生成maps.json的小程序。  
##用这个json，你可以将录像中的地图标识符匹配成客户端中显示的名称  
##A small tool that convert WoT Client files to maps.json.   
##With this Json file, you can match maps' name in replays to the tanks' name displayed in your cilent!  
  
####maps.json的格式尽可能的符合  
####https://github.com/Phalynx/WoT-Dossier-Cache-to-JSON/blob/master/maps.json  
####The format of maps.json is similar with  
####https://github.com/Phalynx/WoT-Dossier-Cache-to-JSON/blob/master/maps.json  
  
在maps.json中，每一张地图是一个对象，每个对象有3个键  
In tanks.json every tank is an object, every object have 8 keys:  
mapid:整数           地图id  
mapid:int            map identifier  
mapidname:字符串     用于匹配po文件中的翻译，也是地图在录像文件中的标识符  
mapidname:string     used to match translations, also maps' name in replays  
mapname:字符串       地图在客户端中显示的名称，中文客户端，这项就是中文  
mapname:string       maps' name displayed in your cilent.If you are using chinese cilent,you'll get chinese name.  
  
Sorry for my poor english,this program use Chinese as its language.  
If everything works well, you only need to input the path of WoT Cilent like C:\Games\World_of_Tanks, and just press ENTERs.  
When the console reads "完成！ztyzbb于2016.11.03敬上！", it means the process finished well without any critical problem.  
If the console reads "检测到潜在的数据缺失！", it means some maps miss the mapname key!  
If you want to edit the maps.json, you may replace "}," with "},^r^n" in your text editer, then the file will be more readable.  
