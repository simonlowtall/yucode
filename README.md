# 程式撰寫案例
這裡存放了了幾個自己撰寫的程式範例，包含簡報自動產製、資料介接程式及Line notify推播程式，程式皆以Python3撰寫，因為是自學的，肯定寫的不是很好，但用起來滿穩定的，近年來陸陸續續寫了不少，但都大同小異。

** **警告** ** 範例在測試前請拉我放在dockerhub上的image run(不會使用docker的抱歉了..可以google一下怎麼使用)，因為不少程式有用到非開放API，不方便放上來，就作了幾個假的API供測試，拉的位置在下面。

>docker pull simonlowtall/yutestapi:latest

## 1. 中央災害應變中心即時開設
* 由中央災害應變中心opendata抓取開設資訊匯入資料庫，內含sql檔可供測試。
* 資料庫位置及帳號密碼在setting.xml中設定
## 2. 彰化縣政府雨量站介接
* 介接彰化縣政府雨量API資料，涉及授權驗證(原始碼24行設定)，所以測試者需要有授權才可使用。
* 內含sql檔可供測試。
* 資料庫位置及帳號密碼在setting.xml中設定。
## 3. 抽水站警戒推播程式
* 抓取API中水位資料，經過判斷後利用Line Notify推播相關警戒資訊。
* 同級的警戒，可在setting.xml中設定重複推播頻率，原則上警戒提升一定推播，下降則不推播。
* 由warninginfo.xml設定警戒水位及警戒文字。
* 由logfile.log儲存狀態。
## 4. 接水資源物聯網資料
* 由水資源物聯網IOW抓取相關資訊。
* 資料庫位置、帳號密碼及IOW url、client_id、client_secret等在setting.xml中設定。
* 因資料表涉及私人公司原創的設計，因此不提供sql檔測試，僅供參考。
## 5. 自動簡報程式
* 從API、網站中抓取資訊、圖片自動產生簡報
* 使用到python-pptx及selenium爬蟲，爬蟲是爬公司自己開發的網站，因為懶得再製圖XD，直接用網站產好的圖比較快。
* chromedriver.exe測試時記得到網路上找符合自己瀏覽器的版本更新。
## 6. 觀測雨量介接氣象局
* 由氣象局opendata抓雨量資料匯入資料庫。
* 第27行Authorization的token是我自己的，可以的話請測試者去氣象局opendata申請一個。
* 內含sql檔可供測試。
* 資料庫位置及帳號密碼在setting.xml中設定
