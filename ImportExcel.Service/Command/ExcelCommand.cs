using ImportExcel.BLL.Models;
using ImportExcel.Service.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace ImportExcel.Service.Command
{
    public class ExcelCommand : IExcelCommand
    {
        public ServiceResult Execute()
        {
            return new ServiceResult();
        }

        public DataTable ExecuteDataTable(DataTable dt, string conn)
        {
            using (OleDbConnection connection = new OleDbConnection(conn))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Siparişler$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    int i = 1;
                    while (dr.Read())
                    {
                        i++;
                        var orderId = dr[0];
                        var orderDate = dr[1];
                        var userId = dr[2];
                        var name = dr[3];
                        var lastName = dr[4];
                        var email = dr[5];
                        var shipMobilepPhone = dr[6];
                        var orderStatus = dr[7];
                        var shipAdressId = dr[8];
                        var bilAddressId = dr[9];
                        var shipperId = dr[10];
                        var paymentmethodSystemName = dr[11];
                        var cartBankName = dr[12];
                        var cardNumber = dr[13];
                        var cardOwner = dr[14]; ////
                        var shippingCost = dr[15];
                        var interestCost = dr[16];
                        var installMeenCount = dr[17];
                        var totalAmount = dr[18];
                        var paymentCurrency = dr[19];
                        var userRole = dr[20];
                        var ipAddress = dr[21];

                        int orderStatusId = 0;
                        if (orderStatus.ToString() == "Teslim Edildi")
                            orderStatusId = 30;
                        else if (orderStatus.ToString() == "İade Edildi")
                            orderStatusId = 50;
                        else if (orderStatus.ToString() == "İptal Edildi")
                            orderStatusId = 40;
                        else if (orderStatus.ToString() == "Kısmi İade Edildi")
                            orderStatusId = 45;

                        string SearchByColumn = "Id=" + orderId;
                        DataRow[] hasRows = dt.Select(SearchByColumn);

                        if (!string.IsNullOrEmpty(orderId.ToString()) && userRole != null /*&& userRole.ToString() != "Guest"*/ && hasRows.Length == 0)
                        {
                            DataRow row = dt.NewRow();
                            var guidID = Guid.NewGuid();
                            row["OrderGuid"] = guidID;
                            ///
                            row["Id"] = orderId;
                            row["StoreId"] = 1;
                            row["CustomerId"] = userId;
                            row["BillingAddressId"] = bilAddressId;
                            row["ShippingAddressId"] = shipAdressId;
                            row["PickUpInStore"] = false;
                            row["OrderStatusId"] = orderStatusId;
                            row["ShippingStatusId"] = 35;
                            row["PaymentStatusId"] = 30;
                            row["PaymentMethodSystemName"] = paymentmethodSystemName;
                            row["CustomerCurrencyCode"] = paymentCurrency;
                            row["CurrencyRate"] = 1.0;
                            row["CustomerTaxDisplayTypeId"] = 0;
                            row["VatNumber"] = null;
                            row["OrderSubtotalInclTax"] = 0.0;
                            row["OrderSubtotalExclTax"] = 0.0;
                            row["OrderSubTotalDiscountInclTax"] = 0.0;
                            row["OrderSubTotalDiscountExclTax"] = 0.0;
                            row["OrderShippingInclTax"] = 0.0;
                            row["OrderShippingExclTax"] = 0.0;
                            row["PaymentMethodAdditionalFeeInclTax"] = 0.0;
                            row["PaymentMethodAdditionalFeeExclTax"] = 0.0;
                            row["OrderTax"] = 0.0;
                            row["OrderDiscount"] = 0.0;
                            row["OrderTotal"] = totalAmount;
                            row["RefundedAmount"] = 0.0;
                            row["RewardPointsWereAdded"] = 0.0;
                            row["CheckoutAttributeDescription"] = "";
                            row["CheckoutAttributesXml"] = null;
                            row["CustomerLanguageId"] = 2;
                            row["AffiliateId"] = 0;
                            row["CustomerIp"] = ipAddress;
                            row["AllowStoringCreditCardNumber"] = 0;
                            row["CardType"] = null;
                            row["CardName"] = cartBankName;
                            row["CardNumber"] = cardNumber;
                            row["MaskedCreditCardNumber"] = cardNumber;
                            row["CardCvv2"] = null;
                            row["CardExpirationMonth"] = null;
                            row["CardExpirationYear"] = null;
                            row["PaymentCardName"] = null;
                            row["PaymentBankName"] = cartBankName;
                            row["PaymentGatewayName"] = "";
                            row["PaymentInstallmentCount"] = installMeenCount;
                            row["PaymentTypeErpCode"] = null;
                            row["AuthorizationTransactionId"] = null;
                            row["AuthorizationTransactionCode"] = null;
                            row["AuthorizationTransactionResult"] = null;
                            row["CaptureTransactionId"] = null;
                            row["CaptureTransactionResult"] = null;
                            row["SubscriptionTransactionId"] = null;
                            row["PaidDateUtc"] = DateTime.Now;
                            row["ShippingRateComputationMethodSystemName"] = "";
                            row["CustomValuesXml"] = null;
                            row["Deleted"] = false;
                            row["CreatedOnUtc"] = orderDate;
                            row["BukoliInvoiceNo"] = null;
                            row["FraudStatusId"] = 0;
                            row["OrderNote"] = null;
                            row["GiftNote"] = null;
                            row["IsExported"] = true;
                            row["ExportedOnUtc"] = DateTime.Now;
                            row["ErpCode"] = null;
                            row["BillingNumber"] = null;
                            row["OrderTypeId"] = 0;
                            row["Is3DOrder"] = 0;
                            row["IsMobileOrder"] = 0;
                            row["RelatedReferral"] = null;
                            row["OrderNumber"] = orderId;
                            row["ReturnOrderTrackingNumber"] = null;
                            row["LoyaltyCardNumber"] = null;
                            row["SMSVerificationCode"] = null;
                            row["IsImpersonated"] = false;
                            row["ImpersonatedById"] = DBNull.Value;
                            row["ImpersonatedStoreId"] = DBNull.Value;
                            row["ImpersonatedStaffId"] = DBNull.Value;
                            row["PaymentMethodAdditionalFeePercent"] = 0.0;
                            row["ReturnedOnUtc"] = DBNull.Value;
                            row["PhysicalStorePosDeviceId"] = DBNull.Value;
                            row["ErpStatusId"] = DBNull.Value;
                            row["ErpTransactionCode"] = DBNull.Value;
                            row["ErpTransactionDate"] = DBNull.Value;
                            row["ErpInvoiceNumber"] = DBNull.Value;
                            row["ErpCode2"] = DBNull.Value;
                            row["OrderErpInformation"] = DBNull.Value;
                            row["OrderFreight"] = 0.0;
                            row["UsedLoyaltyPointTotal"] = 0.0;
                            row["PaymentCurrencyCode"] = paymentCurrency;
                            row["CargoLineId"] = DBNull.Value;
                            row["ReturnExportedOnUtc"] = DBNull.Value;
                            row["ReturnOrderErpInformation"] = DBNull.Value;
                            row["CheckOutSessionValue"] = DBNull.Value;
                            row["PaymentMethodUsedPoint"] = 0.0;
                            row["UsedLoyaltyPointMultiplir"] = 1.0;
                            row["InvoiceUrl"] = DBNull.Value;
                            row["DeliveryBarcode"] = DBNull.Value;
                            row["ShippingMethod"] = shipperId;
                            row["UsedLoyaltyPointMultiplier"] = 0.0;

                            dt.Rows.Add(row);
                        }
                    }
                }
            }

            return dt;
        }
    }
}
