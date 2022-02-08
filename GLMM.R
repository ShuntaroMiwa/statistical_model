library(openxlsx)
library(Matrix)
library(lme4)
library(car)
library(lattice)
library(plm)
library(texreg)
library(nlme)
library(dplyr)
#install.packages("performance")
library(predictmeans)
library(ggplot2)
library(performance)

#線グラフ化（各個人）
excl <- read.xlsx("元テ???ータファイル1.xlsx")

day <-c(0,3,7,10,14)
plot(x=day, y=excl[1,3:7],type="b",ylim=c(110,160),col="black",xlab="日数",ylab="血圧")
for(i in 2:9){
  points(x=day,y=excl[i,3:7],type="b",
         col=ifelse(excl[i,2]=="対照群","black",ifelse(excl[i,2]=="治療A","red","blue")))}

#線グラフ化（群平均）
y0 <- tapply(excl[,3],excl$Group,mean); y3 <- tapply(excl[,4],excl$Group,mean); y7 <- tapply(excl[,5],excl$Group,mean); y10 <- tapply(excl[,6],excl$Group,mean); y14 <- tapply(excl[,7],excl$Group,mean)
ym<-t(rbind(y0,y3,y7,y10,y14))
cols=ifelse(rownames(ym)[i]=="対照群","black",ifelse(rownames(ym)[i]=="治療A","red","blue"))

plot(x=day,y= ym[1,1:5],type="b",ylim=c(110,160),col="black",xlab="日数",ylab="血圧",lwd=3,cex.lab=1.3)
for(i in 2:3){
  points(x=day, y= ym[i,1:5],type="b",col=cols,lwd=3)
}

labels <- rownames(ym)
legend("topleft", legend = labels, col = cols)

#線グラフ化（標準誤差のエラーバー付き）
sd0 <- tapply(excl[,3],excl$Group,sd); sd3 <- tapply(excl[,4],excl$Group,sd); sd7 <- tapply(excl[,5],excl$Group,sd)
sd10 <- tapply(excl[,6],excl$Group,sd); sd14 <- tapply(excl[,7],excl$Group,sd)
sem<-t(rbind(sd0/sqrt(3),sd3/sqrt(3),sd7/sqrt(3),sd10/sqrt(3),sd14/sqrt(3)))
plot(x=day, y=ym[1,1:5],type="b",ylim=c(110,160),col="red",xlab="日数",ylab="血圧",lwd=3,cex.lab=1.3,pch=15)
arrows(x0=day,y0=ym[1,1:5]-sem[1,1:5], x1=day, y1=ym[1,1:5]+sem[1,1:5], lwd=2,code=3,angle=90,length = 0.05,col="red")
points(x=day+0.15, y=ym[2,],type="b",col="blue",lwd=3,pch=15)
arrows(x0=day+0.15,y0=ym[2,]-sem[2,],x1=day+0.15,y1=ym[2,]+sem[2,],lwd=2,code=3,angle=90,length=0.05,col="blue")
points(x=day+0.3,y=ym[3,1:5],type="b",col="black",lwd=3,pch=15)
arrows(x0=day+0.3,y0=ym[3,]-sem[3,],x1=day+0.3,y1=ym[3,]+sem[3,],lwd=2,code=3,angle=90,length=0.05,col="black")

labels <- rownames(ym)
legend("topleft", legend = labels, col = c("black","red","blue"))

#経時繰り返し測定anova
d <- read.csv("Lanova.csv")
(Ymmm=mean(d$Y))
#(Yimm=tapply(d$Y,d$ID,mean)) # これでは不具合が生じることを確認
########################################
# tapply関数を使う際の自動割り当ての順番を変更する
########################################
d$ID <- factor(d$ID,levels=c("C1","C2","C3","A1","A2","A3","B1","B2","B3"))
d$Group <-factor(d$Group,levels=c("対照","治療A","治療B")) 
########################################
(Yimm=tapply(d$Y,d$ID,mean))#個人(I)の平均
(Ymjm=tapply(d$Y,d$Group,mean))#群(J)の平均
(Ymmk=tapply(d$Y,d$day,mean))#時間(K)の平均

I=9; J=3; K=5; Kv=3:7
Smjk=matrix(0,ncol=K,nrow=J)#3*5（群×時間）
nmjk=matrix(0,ncol=K,nrow=J)#3*5
for(l in 1:nrow(d)){
  if(d$day[l]==0){ k=1}
  if(d$day[l]==3){ k=2}
  if(d$day[l]==7){ k=3}
  if(d$day[l]==10){k=4}
  if(d$day[l]==14){k=5}
  if(d$Group[l]=="対照"){j=1}
  if(d$Group[l]=="治療A"){j=2}
  if(d$Group[l]=="治療B"){j=3}
  Smjk[j,k]=Smjk[j,k]+d$Y[l]#群別・時間別の血圧（合計）
  nmjk[j,k]=nmjk[j,k]+1#群別・時間別の人数のカウント
}
Ymjk=Smjk/nmjk#群別・時間別の血圧（平均）

######################################## # 次に各SSを算出 ########################################
SS_G=0;#群間変動
SS_SG=0;#
SS_T=0;#時間変動
SS_GT=0; #群間*時間変動
SS_ERR=0; #誤差変動
SS_Total=0#総変動
for(l in 1:nrow(d)){
  if(d$ID[l]=="C1"){i=1}
  if(d$ID[l]=="C2"){i=2}
  if(d$ID[l]=="C3"){i=3}
  if(d$ID[l]=="A1"){i=4}
  if(d$ID[l]=="A2"){i=5}
  if(d$ID[l]=="A3"){i=6}
  if(d$ID[l]=="B1"){i=7}
  if(d$ID[l]=="B2"){i=8}
  if(d$ID[l]=="B3"){i=9}
  if(d$Group[l]=="対照"){ j=1} # 添え字j（群）の番号付け
  if(d$Group[l]=="治療A"){j=2}
  if(d$Group[l]=="治療B"){j=3}
  if(d$day[l]==0){k=1} # 添え字k（時間）の番号付け
  if(d$day[l]==3){k=2}
  if(d$day[l]==7){k=3}
  if(d$day[l]==10){k=4}
  if(d$day[l]==14){k=5}
  SS_G=SS_G+(Ymjm[j]-Ymmm)^2#群間変動=(群別・時間平均-全平均)^2
  SS_SG=SS_SG+(Yimm[i]-Ymjm[j])^2#個人・群間変動=(個人平均-群平均)^2
  SS_T=SS_T+ (Ymmk[k]-Ymmm)^2#時間変動=(時間平均-全平均)^2
  SS_GT=SS_GT+ (Ymjk[j,k] - Ymjm[j] - Ymmk[k] + Ymmm)^2#群間*時間変動=(群別・時間別平均-群平均-時間別平均+全平均)^2
  SS_ERR=SS_ERR+ (d$Y[l] - Ymjk[j,k] - Yimm[i] + Ymjm[j])^2#誤差変動=(個人-群別・時間別平均-個人平均+群平均)^2
  SS_Total=SS_Total+(d$Y[l]-Ymmm)^2
}
SS_G
SS_SG
SS_T
SS_GT
SS_ERR
c(SS_Total,SS_G+SS_SG+SS_T+SS_GT+SS_ERR)

#F検定
1-pf(10.032,df1=2,df2=6)
1-pf(10.030,df1=4,df2=24)
1-pf(2.948,df1=8,df2=24)

#混合効果モデル
Lanova <- read.csv("Lanova.csv")
day <- Lanova$day
Y <- Lanova$Y
ID<- Lanova$ID
Group<- Lanova$Group

mode(Y)

a <- lmer(Y?Group+factor(day)+Group:factor(day) +(1|ID), data = Lanova)
summary(a)
AIC(a)
BIC(a)
anova(a)
CookD(a)
plot(a)
r2(a)
print(a, correlation=TRUE)

anova(a,b,c, refit = FALSE)

a1 <- lmer(Y?Group+factor(day)+Group:factor(day) +(Group|ID), data = Lanova)

?isSingular
isSingular(a1, tol = 1e-4)
summary(a1)
anova(a1)
anova(a,a1,refit = FALSE)

p <- ggplot(Lanova, aes(x = day, y = Y, colour = ID)) +
  geom_point(size=3) +
  geom_line(aes(y = predict(a)),size=1) 
print(p)



#混合効果モデルの多重共線性の確認
lm1 <- lm(Y?Group+factor(day)+Group:factor(day), data = Lanova) # 同じ固定効果を OLS で推定
vif(lm1)

plot(a)

qqmath(a)
ranef2 <- ranef(a)#ランダム効果
ranef2
plot(ranef2)
qqmath(ranef2)

#二元配置anova
b <- lm(Y?Group+factor(day)+Group:factor(day), data = Lanova)#群、時間、群・時間の交互作用
summary(b)
anova(b)
AIC(b)
BIC(b)
r2(b)

c <- lm(Y?ID+factor(day), data = Lanova)#個人と時間の効果のみ
summary(c)
anova(c)
AIC(c)
BIC(c)
r2(c)
co1 <- check_collinearity(a)
