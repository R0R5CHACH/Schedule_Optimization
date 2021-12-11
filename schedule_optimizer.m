%Enployee Scheduling Optimization
%Perkins, Dallas
%MAE 509, Arizona State University
clear;

%% Collect Data
data=importdata('E.xlsx');
C{1}=data.data.Coverage(1:14,2:8);
C{2}=data.data.Coverage(17:30,2:8);
C{3}=data.data.Coverage(33:46,2:8);
C{4}=data.data.Coverage(49:62,2:8);
E{1,1}=data.data.E1(1:14,2:8);
E{1,2}=data.data.E1([1:3,5:6],11);
E{2,1}=data.data.E2(1:14,2:8);
E{2,2}=data.data.E2([1:3,5:6],11);
E{3,1}=data.data.E3(1:14,2:8);
E{3,2}=data.data.E3([1:3,5:6],11);
E{4,1}=data.data.E4(1:14,2:8);
E{4,2}=data.data.E4([1:3,5:6],11);
E{5,1}=data.data.E5(1:14,2:8);
E{5,2}=data.data.E5([1:3,5:6],11);
E{6,1}=data.data.E6(1:14,2:8);
E{6,2}=data.data.E6([1:3,5:6],11);
E{7,1}=data.data.E7(1:14,2:8);
E{7,2}=data.data.E7([1:3,5:6],11);
E{8,1}=data.data.E8(1:14,2:8);
E{8,2}=data.data.E8([1:3,5:6],11);
E{9,1}=data.data.E9(1:14,2:8);
E{9,2}=data.data.E9([1:3,5:6],11);
E{10,1}=data.data.E10(1:14,2:8);
E{10,2}=data.data.E10([1:3,5:6],11);
E{11,1}=data.data.E10(1:14,2:8);
E{11,2}=data.data.E10([1:3,5:6],11);
E{12,1}=data.data.E10(1:14,2:8);
E{12,2}=data.data.E10([1:3,5:6],11);
E{13,1}=data.data.E10(1:14,2:8);
E{13,2}=data.data.E10([1:3,5:6],11);
E{14,1}=data.data.E10(1:14,2:8);
E{14,2}=data.data.E10([1:3,5:6],11);
E{15,1}=data.data.E10(1:14,2:8);
E{15,2}=data.data.E10([1:3,5:6],11);
[h,d]=size(C{1}); %[hours, days] business is open
n=15; %number of employees
zh=ones(1,h); zd=ones(d,1); %useful summation vectors

%% Problem Structure
%define variables
C0=zeros(h,d); C1=C0; C2=C0; C3=C0; %matrices for coverage constraints
Sal=[]; Div=[]; F=[]; %initialize objectives and constraints
for i=1:n
    S{i}=binvar(14,7,'full'); %employee shifts
    C0=C0+S{i}; %general coverage check
    C1=C1+E{i,2}(1)*S{i}; %check: employee i's ability to cover system 1
    C2=C2+E{i,2}(2)*S{i};
    C3=C3+E{i,2}(3)*S{i};
    Sal=[Sal E{i,2}(4)*zh*S{i}*zd]; %salary cost function
    Div=[Div zh*(S{i}.*E{i,1})*zd]; %divergence from employee preference
    F=[F; zh*S{i}*zd<=E{i,2}(5)]; %max weekly hours constraint
end
for i=1:h
    for j=1:d
        F=[F; C0(i,j)>=C{1}(i,j)]; %general coverage constraint
        F=[F; C1(i,j)>=C{2}(i,j)]; %system 1 coverage constraint
        F=[F; C2(i,j)>=C{3}(i,j)];
        F=[F; C3(i,j)>=C{4}(i,j)];
    end
end
%split-shift constraints
for k=1:n
    for i=2:(h-1)
        for j=1:d
            F=[F; implies(and(S{k}(i,j)<=0,sum(S{k}(1:i-1,j))>=1),1-S{k}(i+1,j))];
        end
    end
end

%% Solve
opt=sdpsettings('verbose',0,'solver','gurobi');
sol=optimize(F,.6*sum(Sal)+.4*sum(Div),opt);
if sol.problem~=0
    disp('Unsuccessful; problem is poorly constrained. Try relaxing constraints.')
else
    writematrix(value(S{1}),'S.xlsx','sheet','E1','range','B2');
    writematrix(value(S{2}),'S.xlsx','sheet','E2','range','B2');
    writematrix(value(S{3}),'S.xlsx','sheet','E3','range','B2');
    writematrix(value(S{4}),'S.xlsx','sheet','E4','range','B2');
    writematrix(value(S{5}),'S.xlsx','sheet','E5','range','B2');
    writematrix(value(S{6}),'S.xlsx','sheet','E6','range','B2');
    writematrix(value(S{7}),'S.xlsx','sheet','E7','range','B2');
    writematrix(value(S{8}),'S.xlsx','sheet','E8','range','B2');
    writematrix(value(S{9}),'S.xlsx','sheet','E9','range','B2');
    writematrix(value(S{10}),'S.xlsx','sheet','E10','range','B2');
    writematrix(value(S{11}),'S.xlsx','sheet','E11','range','B2');
    writematrix(value(S{12}),'S.xlsx','sheet','E12','range','B2');
    writematrix(value(S{13}),'S.xlsx','sheet','E13','range','B2');
    writematrix(value(S{14}),'S.xlsx','sheet','E14','range','B2');
    writematrix(value(S{15}),'S.xlsx','sheet','E15','range','B2');
    writematrix(value(C0),'S.xlsx','sheet','Coverage','range','B2');
    writematrix(value(C1),'S.xlsx','sheet','Coverage','range','B18');
    writematrix(value(C2),'S.xlsx','sheet','Coverage','range','B34');
    writematrix(value(C3),'S.xlsx','sheet','Coverage','range','B50');
    disp('Optimal scheduling found; schedules written to spreadsheet.')
end