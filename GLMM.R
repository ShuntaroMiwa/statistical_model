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

#���O���t���i�e�l�j
excl <- read.xlsx("���e???�[�^�t�@�C��1.xlsx")

day <-c(0,3,7,10,14)
plot(x=day, y=excl[1,3:7],type="b",ylim=c(110,160),col="black",xlab="����",ylab="����")
for(i in 2:9){
  points(x=day,y=excl[i,3:7],type="b",
         col=ifelse(excl[i,2]=="�ΏƌQ","black",ifelse(excl[i,2]=="����A","red","blue")))}

#���O���t���i�Q���ρj
y0 <- tapply(excl[,3],excl$Group,mean); y3 <- tapply(excl[,4],excl$Group,mean); y7 <- tapply(excl[,5],excl$Group,mean); y10 <- tapply(excl[,6],excl$Group,mean); y14 <- tapply(excl[,7],excl$Group,mean)
ym<-t(rbind(y0,y3,y7,y10,y14))
cols=ifelse(rownames(ym)[i]=="�ΏƌQ","black",ifelse(rownames(ym)[i]=="����A","red","blue"))

plot(x=day,y= ym[1,1:5],type="b",ylim=c(110,160),col="black",xlab="����",ylab="����",lwd=3,cex.lab=1.3)
for(i in 2:3){
  points(x=day, y= ym[i,1:5],type="b",col=cols,lwd=3)
}

labels <- rownames(ym)
legend("topleft", legend = labels, col = cols)

#���O���t���i�W���덷�̃G���[�o�[�t���j
sd0 <- tapply(excl[,3],excl$Group,sd); sd3 <- tapply(excl[,4],excl$Group,sd); sd7 <- tapply(excl[,5],excl$Group,sd)
sd10 <- tapply(excl[,6],excl$Group,sd); sd14 <- tapply(excl[,7],excl$Group,sd)
sem<-t(rbind(sd0/sqrt(3),sd3/sqrt(3),sd7/sqrt(3),sd10/sqrt(3),sd14/sqrt(3)))
plot(x=day, y=ym[1,1:5],type="b",ylim=c(110,160),col="red",xlab="����",ylab="����",lwd=3,cex.lab=1.3,pch=15)
arrows(x0=day,y0=ym[1,1:5]-sem[1,1:5], x1=day, y1=ym[1,1:5]+sem[1,1:5], lwd=2,code=3,angle=90,length = 0.05,col="red")
points(x=day+0.15, y=ym[2,],type="b",col="blue",lwd=3,pch=15)
arrows(x0=day+0.15,y0=ym[2,]-sem[2,],x1=day+0.15,y1=ym[2,]+sem[2,],lwd=2,code=3,angle=90,length=0.05,col="blue")
points(x=day+0.3,y=ym[3,1:5],type="b",col="black",lwd=3,pch=15)
arrows(x0=day+0.3,y0=ym[3,]-sem[3,],x1=day+0.3,y1=ym[3,]+sem[3,],lwd=2,code=3,angle=90,length=0.05,col="black")

labels <- rownames(ym)
legend("topleft", legend = labels, col = c("black","red","blue"))

#�o���J��Ԃ�����anova
d <- read.csv("Lanova.csv")
(Ymmm=mean(d$Y))
#(Yimm=tapply(d$Y,d$ID,mean)) # ����ł͕s��������邱�Ƃ��m�F
########################################
# tapply�֐����g���ۂ̎������蓖�Ă̏��Ԃ�ύX����
########################################
d$ID <- factor(d$ID,levels=c("C1","C2","C3","A1","A2","A3","B1","B2","B3"))
d$Group <-factor(d$Group,levels=c("�Ώ�","����A","����B")) 
########################################
(Yimm=tapply(d$Y,d$ID,mean))#�l(I)�̕���
(Ymjm=tapply(d$Y,d$Group,mean))#�Q(J)�̕���
(Ymmk=tapply(d$Y,d$day,mean))#����(K)�̕���

I=9; J=3; K=5; Kv=3:7
Smjk=matrix(0,ncol=K,nrow=J)#3*5�i�Q�~���ԁj
nmjk=matrix(0,ncol=K,nrow=J)#3*5
for(l in 1:nrow(d)){
  if(d$day[l]==0){ k=1}
  if(d$day[l]==3){ k=2}
  if(d$day[l]==7){ k=3}
  if(d$day[l]==10){k=4}
  if(d$day[l]==14){k=5}
  if(d$Group[l]=="�Ώ�"){j=1}
  if(d$Group[l]=="����A"){j=2}
  if(d$Group[l]=="����B"){j=3}
  Smjk[j,k]=Smjk[j,k]+d$Y[l]#�Q�ʁE���ԕʂ̌����i���v�j
  nmjk[j,k]=nmjk[j,k]+1#�Q�ʁE���ԕʂ̐l���̃J�E���g
}
Ymjk=Smjk/nmjk#�Q�ʁE���ԕʂ̌����i���ρj

######################################## # ���ɊeSS���Z�o ########################################
SS_G=0;#�Q�ԕϓ�
SS_SG=0;#
SS_T=0;#���ԕϓ�
SS_GT=0; #�Q��*���ԕϓ�
SS_ERR=0; #�덷�ϓ�
SS_Total=0#���ϓ�
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
  if(d$Group[l]=="�Ώ�"){ j=1} # �Y����j�i�Q�j�̔ԍ��t��
  if(d$Group[l]=="����A"){j=2}
  if(d$Group[l]=="����B"){j=3}
  if(d$day[l]==0){k=1} # �Y����k�i���ԁj�̔ԍ��t��
  if(d$day[l]==3){k=2}
  if(d$day[l]==7){k=3}
  if(d$day[l]==10){k=4}
  if(d$day[l]==14){k=5}
  SS_G=SS_G+(Ymjm[j]-Ymmm)^2#�Q�ԕϓ�=(�Q�ʁE���ԕ���-�S����)^2
  SS_SG=SS_SG+(Yimm[i]-Ymjm[j])^2#�l�E�Q�ԕϓ�=(�l����-�Q����)^2
  SS_T=SS_T+ (Ymmk[k]-Ymmm)^2#���ԕϓ�=(���ԕ���-�S����)^2
  SS_GT=SS_GT+ (Ymjk[j,k] - Ymjm[j] - Ymmk[k] + Ymmm)^2#�Q��*���ԕϓ�=(�Q�ʁE���ԕʕ���-�Q����-���ԕʕ���+�S����)^2
  SS_ERR=SS_ERR+ (d$Y[l] - Ymjk[j,k] - Yimm[i] + Ymjm[j])^2#�덷�ϓ�=(�l-�Q�ʁE���ԕʕ���-�l����+�Q����)^2
  SS_Total=SS_Total+(d$Y[l]-Ymmm)^2
}
SS_G
SS_SG
SS_T
SS_GT
SS_ERR
c(SS_Total,SS_G+SS_SG+SS_T+SS_GT+SS_ERR)

#F����
1-pf(10.032,df1=2,df2=6)
1-pf(10.030,df1=4,df2=24)
1-pf(2.948,df1=8,df2=24)

#�������ʃ��f��
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



#�������ʃ��f���̑��d�������̊m�F
lm1 <- lm(Y?Group+factor(day)+Group:factor(day), data = Lanova) # �����Œ���ʂ� OLS �Ő���
vif(lm1)

plot(a)

qqmath(a)
ranef2 <- ranef(a)#�����_������
ranef2
plot(ranef2)
qqmath(ranef2)

#�񌳔z�uanova
b <- lm(Y?Group+factor(day)+Group:factor(day), data = Lanova)#�Q�A���ԁA�Q�E���Ԃ̌��ݍ�p
summary(b)
anova(b)
AIC(b)
BIC(b)
r2(b)

c <- lm(Y?ID+factor(day), data = Lanova)#�l�Ǝ��Ԃ̌��ʂ̂�
summary(c)
anova(c)
AIC(c)
BIC(c)
r2(c)
co1 <- check_collinearity(a)