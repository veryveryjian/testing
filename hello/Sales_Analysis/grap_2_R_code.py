# library(ggplot2)
# library(readr)
#
# # 데이터 로드
# data <- read_csv('sale_analysis2.CSV')
#
# # ggplot을 사용하여 바와 라인 차트 생성
# ggplot(data, aes(x=Cabinets_sub_code)) +
#   geom_bar(aes(y=Sum_of_Qty), stat="identity", fill="blue", alpha=0.6) +
#   geom_line(aes(y=Sum_of_CBM * 10), color="red", group=1) + # CBM을 확대하여 가시성 향상
#   scale_y_continuous(
#     name = "Sum of Qty",
#     sec.axis = sec_axis(~./10, name="Sum of CBM") # CBM 스케일을 조정
#   ) +
#   labs(title="Cabinet Sales Analysis") +
#   theme_minimal()
