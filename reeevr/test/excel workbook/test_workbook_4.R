library(reeevr)
library(BCEA)
numberOfRuns = 100000
converter_validate = FALSE
Frontend_E5 = 20000
Frontend_E6 = "Probabilistic"
Inputparameters_J4 = 10
Inputparameters_K4 = 90
Inputparameters_J5 = 2
Inputparameters_K5 = 78
Inputparameters_J6 = 4
Inputparameters_K6 = 156
Inputparameters_J7 = 30
Inputparameters_K7 = 70
Inputparameters_J8 = 4
Inputparameters_K8 = 76
Inputparameters_J9 = 1
Inputparameters_K9 = 99
Inputparameters_H10 = 10
Inputparameters_H11 = 500
Inputparameters_H13 = 400
Inputparameters_H14 = 2000
Inputparameters_H16 = 12
Inputparameters_I16 = 1
Inputparameters_H17 = 20
Inputparameters_I17 = 2
Inputparameters_H18 = 12
Inputparameters_I18 = 1
Costing_E4 = 5000
Costing_E5 = 400
Costing_E6 = 100
Costing_E7 = 150
Costing_E8 = 100
Costing_E9 = 100
QALYs_D4 = 14
QALYs_D5 = 15
QALYs_D6 = 12
QALYs_D7 = 16
QALYs_D8 = 14
QALYs_D9 = 14
QALYs_D10 = 15
QALYs_D11 = 13
QALYs_D12 = 12
QALYs_D13 = 16
QALYs_D14 = 16
QALYs_D15 = 10
QALYs_D16 = 14
QALYs_D17 = 15
QALYs_D18 = 14
QALYs_D19 = 13
QALYs_D20 = 17
QALYs_D21 = 12
QALYs_D22 = 18
QALYs_D23 = 14
analysis_type = Frontend_E6
willingness_to_pay = Frontend_E5
Inputparameters_G4 = qbeta(excel_rand(numberOfRuns, converter_validate),Inputparameters_J4,Inputparameters_K4)
Inputparameters_H4 = Inputparameters_J4 / ( Inputparameters_J4 + Inputparameters_K4 )
Inputparameters_G5 = qbeta(excel_rand(numberOfRuns, converter_validate),Inputparameters_J5,Inputparameters_K5)
Inputparameters_H5 = Inputparameters_J5 / ( Inputparameters_J5 + Inputparameters_K5 )
Inputparameters_G6 = qbeta(excel_rand(numberOfRuns, converter_validate),Inputparameters_J6,Inputparameters_K6)
Inputparameters_H6 = Inputparameters_J6 / ( Inputparameters_J6 + Inputparameters_K6 )
Inputparameters_G7 = qbeta(excel_rand(numberOfRuns, converter_validate),Inputparameters_J7,Inputparameters_K7)
Inputparameters_H7 = Inputparameters_J7 / ( Inputparameters_J7 + Inputparameters_K7 )
Inputparameters_H8 = Inputparameters_J8 / ( Inputparameters_J8 + Inputparameters_K8 )
Inputparameters_H9 = Inputparameters_J9 / ( Inputparameters_J9 + Inputparameters_K9 )
Inputparameters_F10 = Inputparameters_H10
Inputparameters_G10 = Inputparameters_F10
Inputparameters_F11 = Inputparameters_H11
Inputparameters_G11 = Inputparameters_F11
Inputparameters_H12 = excel_sum(list(Costing_E4, Costing_E5, Costing_E6, Costing_E7, Costing_E8, Costing_E9))
Inputparameters_F13 = Inputparameters_H13
Inputparameters_G13 = Inputparameters_F13
Inputparameters_F14 = Inputparameters_H14
Inputparameters_G14 = Inputparameters_F14
Inputparameters_H15 = excel_average(list(QALYs_D4, QALYs_D5, QALYs_D6, QALYs_D7, QALYs_D8, QALYs_D9, QALYs_D10, QALYs_D11, QALYs_D12, QALYs_D13, QALYs_D14, QALYs_D15, QALYs_D16, QALYs_D17, QALYs_D18, QALYs_D19, QALYs_D20, QALYs_D21, QALYs_D22, QALYs_D23))
Inputparameters_I15 = excel__xlfn.stdev.s(list(QALYs_D4, QALYs_D5, QALYs_D6, QALYs_D7, QALYs_D8, QALYs_D9, QALYs_D10, QALYs_D11, QALYs_D12, QALYs_D13, QALYs_D14, QALYs_D15, QALYs_D16, QALYs_D17, QALYs_D18, QALYs_D19, QALYs_D20, QALYs_D21, QALYs_D22, QALYs_D23)) / sqrt(20)
Inputparameters_F16 = Inputparameters_H16
Inputparameters_G16 = qnorm(excel_rand(numberOfRuns, converter_validate),Inputparameters_H16,Inputparameters_I16)
Inputparameters_F17 = Inputparameters_H17
Inputparameters_G17 = qnorm(excel_rand(numberOfRuns, converter_validate),Inputparameters_H17,Inputparameters_I17)
Inputparameters_F18 = Inputparameters_H18
Inputparameters_G18 = qnorm(excel_rand(numberOfRuns, converter_validate),Inputparameters_H18,Inputparameters_I18)
Inputparameters_F4 = Inputparameters_H4
Inputparameters_F5 = Inputparameters_H5
Inputparameters_F6 = Inputparameters_H6
Inputparameters_F7 = Inputparameters_H7
Inputparameters_F8 = Inputparameters_H8
Inputparameters_G8 = Inputparameters_F8
Inputparameters_F9 = Inputparameters_H9
Inputparameters_G9 = Inputparameters_F9
Inputparameters_E10_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F10)
    }
    else{
        return(Inputparameters_G10)
    }
}
Inputparameters_E10 = Inputparameters_E10_if_1_1()
Inputparameters_E11_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F11)
    }
    else{
        return(Inputparameters_G11)
    }
}
Inputparameters_E11 = Inputparameters_E11_if_1_1()
Inputparameters_F12 = Inputparameters_H12
Inputparameters_G12 = Inputparameters_F12
Inputparameters_E13_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F13)
    }
    else{
        return(Inputparameters_G13)
    }
}
Inputparameters_E13 = Inputparameters_E13_if_1_1()
Inputparameters_E14_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F14)
    }
    else{
        return(Inputparameters_G14)
    }
}
Inputparameters_E14 = Inputparameters_E14_if_1_1()
Inputparameters_F15 = Inputparameters_H15
Inputparameters_G15 = qnorm(excel_rand(numberOfRuns, converter_validate),Inputparameters_H15,Inputparameters_I15)
Inputparameters_E16_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F16)
    }
    else{
        return(Inputparameters_G16)
    }
}
Inputparameters_E16 = Inputparameters_E16_if_1_1()
Inputparameters_E17_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F17)
    }
    else{
        return(Inputparameters_G17)
    }
}
Inputparameters_E17 = Inputparameters_E17_if_1_1()
Inputparameters_E18_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F18)
    }
    else{
        return(Inputparameters_G18)
    }
}
Inputparameters_E18 = Inputparameters_E18_if_1_1()
Inputparameters_E4_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F4)
    }
    else{
        return(Inputparameters_G4)
    }
}
Inputparameters_E4 = Inputparameters_E4_if_1_1()
Inputparameters_E5_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F5)
    }
    else{
        return(Inputparameters_G5)
    }
}
Inputparameters_E5 = Inputparameters_E5_if_1_1()
Inputparameters_E6_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F6)
    }
    else{
        return(Inputparameters_G6)
    }
}
Inputparameters_E6 = Inputparameters_E6_if_1_1()
Inputparameters_E7_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F7)
    }
    else{
        return(Inputparameters_G7)
    }
}
Inputparameters_E7 = Inputparameters_E7_if_1_1()
Inputparameters_E8_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F8)
    }
    else{
        return(Inputparameters_G8)
    }
}
Inputparameters_E8 = Inputparameters_E8_if_1_1()
Inputparameters_E9_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F9)
    }
    else{
        return(Inputparameters_G9)
    }
}
Inputparameters_E9 = Inputparameters_E9_if_1_1()
Inputparameters_E12_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F12)
    }
    else{
        return(Inputparameters_G12)
    }
}
Inputparameters_E12 = Inputparameters_E12_if_1_1()
Inputparameters_E15_if_1_1 <- function(){
    if(analysis_type=="Deterministic"){
        return(Inputparameters_F15)
    }
    else{
        return(Inputparameters_G15)
    }
}
Inputparameters_E15 = Inputparameters_E15_if_1_1()
Engine_E5 = Inputparameters_E10 + Inputparameters_E6 * Inputparameters_E13 + Inputparameters_E4 * Inputparameters_E12 + Inputparameters_E8 * Inputparameters_E14
Engine_F5 = Inputparameters_E11 + Inputparameters_E7 * Inputparameters_E13 + Inputparameters_E5 * Inputparameters_E12 + Inputparameters_E9 + Inputparameters_E14
Engine_E6 = Inputparameters_E6 * Inputparameters_E16 + Inputparameters_E4 * Inputparameters_E15 + ( 1 - Inputparameters_E6 - Inputparameters_E4 - Inputparameters_E8 ) * Inputparameters_E17 + Inputparameters_E8 * Inputparameters_E18
Engine_F6 = Inputparameters_E7 * Inputparameters_E16 + Inputparameters_E5 * Inputparameters_E15 + ( 1 - Inputparameters_E7 - Inputparameters_E5 - Inputparameters_E9 ) * Inputparameters_E17 + Inputparameters_E9 * Inputparameters_E18
Engine_E7 = Engine_E6 * willingness_to_pay - Engine_E5
Engine_F7 = Engine_F6 * willingness_to_pay - Engine_F5
psafile <- 'excel workbook//psa.txt'
PSA_output <- PSA_output(psafile, dec = '.',Engine_E5,Engine_F5,Engine_E6,Engine_F6,Engine_E7,Engine_F7)
