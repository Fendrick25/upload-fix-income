package com.ciptadana.uploadfixincome;

import lombok.Builder;
import lombok.Data;
import lombok.ToString;

@Data
@Builder
@ToString
public class FixIncome {
    private String issuerName;
    private String category;
    private String type;
    private String coupon;
    private String rating;
    private String maturity;
    private String bidPrice;
    private String offerPrice;
    private String yieldBid;
    private String yieldOffer;
    private String currency;
    private String accountMin;
}
