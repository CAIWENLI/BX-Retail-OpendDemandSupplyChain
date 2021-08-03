keepa_isbn_file <- list.files("C:/Users/lisal/OneDrive - bookxchange.com/Retail Reporting/Purchase Analysis/Keepa Files/Long Research Files/2021-02-16-EU", pattern="*.xlsx", full.names=TRUE)
keepa_isbn <- lapply(keepa_isbn_file, read_xlsx, sheet = 1)
for (i in 1:length(keepa_isbn)){keepa_isbn[[i]]<-cbind(keepa_isbn[[i]],keepa_isbn_file[i])}
keepa_isbn_all <- do.call("rbind", keepa_isbn) 
keepa_buy <- keepa_isbn_all
tbl_vars(keepa_buy)
# Data checking and selection
keepa_buy <- keepa_buy %>% 
  select(`Product Codes: EAN`, ASIN, `Title`,`Buy Box: Current`, `Buy Box: 30 days avg.`,`Buy Box: Lowest`,`Used: 30 days avg.`,`Sales Rank: Current`,`Sales Rank: 30 days avg.`,`Sales Rank: 90 days avg.`,`Sales Rank: 180 days avg.`,`Sales Rank: Lowest`,`New Offer Count: Current`, `Used Offer Count: Current`, `Publication Date`,`Buy Box Seller`, `FBA Fees:`, `Sales Rank: Drops last 90 days`, `Buy Box: Is FBA`, `Rental: Current`, `Rental: 30 days avg.`, `Rental: Lowest`)

# keepa_buy$publication_date <- as.Date(keepa_buy$publication_date)
library(stringr)
keepa_buy <- str_split_fixed(keepa_buy$`Product Codes: EAN`, ",", 4) %>% 
  cbind(keepa_buy)

tbl_vars(keepa_buy)

colnames(keepa_buy) <- c("V1","V2","V3","V4","isbn", "asin","title", "buybox_current", "buybox_30day_avg", "buybox_lowest","used_30day_avg","sale_rank_current","sale_rank_30","sale_rank_90","sale_rank_180", "sale_rank_lowest", "new_offer_count","used_offer_count","publication_date", "buybox_seller", "fba_fees", "sale_rank_drop_90days", "buybox_isfba", "rental_current", "rental_30day_avg", "rental_lowest")

keepa_buy <- keepa_buy %>% 
  mutate(isbn_new =  ifelse(startsWith(V2, " 97"), V2, 
                            ifelse(startsWith(V3, " 97"), V3, V1))) %>% 
  select("isbn_new", "isbn", "asin", "title","buybox_current", "buybox_30day_avg", "buybox_lowest","used_30day_avg","sale_rank_current","sale_rank_30","sale_rank_90","sale_rank_180", "sale_rank_lowest","new_offer_count","used_offer_count","publication_date", "buybox_seller", "fba_fees", "sale_rank_drop_90days", "buybox_isfba", "rental_current", "rental_30day_avg", "rental_lowest") 

keepa_buy$buybox_current <- as.numeric(gsub("[\\$,]", "",keepa_buy$buybox_current))
keepa_buy$buybox_30day_avg <- as.numeric(gsub("[\\$,]", "",keepa_buy$buybox_30day_avg))

keepa_buy$isbn_new <- trimws(keepa_buy$isbn_new)

keepa_buy <- keepa_buy %>% 
  arrange(buybox_30day_avg) %>% 
  distinct(isbn_new,.keep_all = TRUE)

keepa_buy[,c(5:15)]<- lapply(keepa_buy[,c(5:15)], as.numeric)
keepa_buy[,c(5:15)][is.na(keepa_buy[,c(5:15)])] <- 0

keepa_buy$isbn_new <- as.character(keepa_buy$isbn_new)
write.xlsx(list("Rental Information" = keepa_buy),paste("C:/Users/lisal/OneDrive - bookxchange.com/Retail Reporting/Purchase Analysis/ISBN_List_TX_", sep = "", today(), ".xlsx"))

