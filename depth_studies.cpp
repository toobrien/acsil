#include "sierrachart.h"
// #include <string>

// For reference, refer to this page:
// https://www.sierrachart.com/index.php?page=doc/AdvancedCustomStudyInterfaceAndLanguage.php

SCDLLName("depth_studies")

SCSFExport scsf_large_orders(SCStudyInterfaceRef sc) {

	#define input_head_col  0
	#define input_val_col	1
	#define spacer_col_a	2
	#define bid_lvls_col	3
	#define bid_qtys_col   	4
	#define bid_tics_col    5
	#define spacer_col_b    6
	#define ask_lvls_col 	7
	#define ask_qtys_col   	8
	#define ask_tics_col	9
	
	#define symbol_row_offset			0
	#define max_outputs_row_offset		1
	#define threshold_row_offset		2
	#define alert_distance_row_offset	3

	#define from_high_row	4
	#define from_low_row	5

	// set defaults

	SCInputRef 	base_row	= sc.Input[0];

	if (sc.SetDefaults) {
		
		sc.GraphName = "large_orders";
		sc.AutoLoop = 0;
		sc.UsesMarketDepthData = 1;
		sc.HideStudy = 1;

		base_row.Name	= "base row";
		base_row.SetInt(-1);

		return;
		
	}

	// initialize inputs

	int base_row_val;
	double d_max_outputs = -1;
	double d_threshold = -1;
	double d_alert_distance = -1;

	base_row_val = base_row.GetInt();

	if (base_row_val < 0)

		// user has not properly initialized the study

		return;
	
	void * h = sc.GetSpreadsheetSheetHandleByName("depth_sheet", "main", false);
	
	sc.GetSheetCellAsDouble(h, input_val_col, base_row_val + max_outputs_row_offset, d_max_outputs);
	sc.GetSheetCellAsDouble(h, input_val_col, base_row_val + threshold_row_offset, d_threshold);
	sc.GetSheetCellAsDouble(h, input_val_col, base_row_val + alert_distance_row_offset, d_alert_distance);

	int max_outputs 	= (int)d_max_outputs;
	int threshold		= (int)d_threshold;
	int alert_distance	= (int)d_alert_distance;

	// sc.AddMessageToLog(("max_outputs: " + std::to_string(max_outputs)).c_str(), 1);
	// sc.AddMessageToLog(("threshold: " + std::to_string(threshold)).c_str(), 1);
	// sc.AddMessageToLog(("alert_distance: " + std::to_string(alert_distance)).c_str(), 1);

	if (threshold <= 0 || max_outputs <= 0)

		// user has not properly initialized the study

		return;

	// record bids and asks where qty > threshold.
	// also find distance from high and low if
	// reasonably close.

	float	bid_lvls[max_outputs];
	float	bid_qtys[max_outputs];
	int   	bid_tics[max_outputs];
	float 	ask_lvls[max_outputs];
	float 	ask_qtys[max_outputs];
	int		ask_tics[max_outputs];

	int len_bids = 0;
	int len_asks = 0;

	int from_high	= -1;
	int from_low	= -1;

	int max_levels = sc.GetMaximumMarketDepthLevels();
	
	if (max_levels > sc.ArraySize)
		
		max_levels = sc.ArraySize;

	int max_bid_levels = min(max_levels, sc.GetBidMarketDepthNumberOfLevels());
	int max_ask_levels = min(max_levels, sc.GetAskMarketDepthNumberOfLevels());

	s_MarketDepthEntry e;

	for (int i = 0; i < max_bid_levels; i++) {

		sc.GetBidMarketDepthEntryAtLevel(e, i);

		if (e.Quantity > threshold) {

			bid_lvls[len_bids] = e.AdjustedPrice;
			bid_qtys[len_bids] = static_cast<float>(e.Quantity);
			bid_tics[len_bids] = i;

			len_bids++;

		}

		if (e.AdjustedPrice == sc.DailyLow)

			from_low = i;

		if (len_bids >= max_outputs) 
		
			break;

	}

	for (int i = 0; i < max_ask_levels; i++) {

		sc.GetAskMarketDepthEntryAtLevel(e, i);

		if (e.Quantity > threshold) {

			ask_lvls[len_asks] = e.AdjustedPrice;
			ask_qtys[len_asks] = static_cast<float>(e.Quantity);
			ask_tics[len_asks] = i;

			len_asks++;

		}

		if (e.AdjustedPrice == sc.DailyHigh)

			from_high = i;

		if (len_asks >= max_outputs)

			break;

	}

	// clear spreadsheet
	
	int lo = base_row_val + 1;
	int hi = lo + max_outputs;
	SCString clr = "x";

	for (int i = lo; i < hi; i++) {

		sc.SetSheetCellAsString(h, bid_lvls_col, i, clr);
		sc.SetSheetCellAsString(h, bid_qtys_col, i, clr);
		sc.SetSheetCellAsString(h, bid_tics_col, i, clr);

	}

	for (int i = lo; i < hi; i++) {

		sc.SetSheetCellAsString(h, ask_lvls_col, i, clr);
		sc.SetSheetCellAsString(h, ask_qtys_col, i, clr);
		sc.SetSheetCellAsString(h, ask_tics_col, i, clr);

	}

	sc.SetSheetCellAsString(h, input_val_col, base_row_val + from_high_row, clr);
	sc.SetSheetCellAsString(h, input_val_col, base_row_val + from_low_row, clr);
	
	// put all bids and asks where qty > threshold into spreadsheet

	hi = lo + len_bids;
	int j = 0;

	for (int i = lo; i < hi; i++) {

		sc.SetSheetCellAsDouble(h, bid_lvls_col, i, bid_lvls[j]);
		sc.SetSheetCellAsDouble(h, bid_qtys_col, i, bid_qtys[j]);
		sc.SetSheetCellAsDouble(h, bid_tics_col, i, bid_tics[j++]);

	}

	hi = lo + len_asks;
	j = 0;

	for (int i = lo; i < hi; i++) {

		sc.SetSheetCellAsDouble(h, ask_lvls_col, i, ask_lvls[j]);
		sc.SetSheetCellAsDouble(h, ask_qtys_col, i, ask_qtys[j]);
		sc.SetSheetCellAsDouble(h, ask_tics_col, i, ask_tics[j++]);

	}

	// set from high / low

	if (from_high > 0)

		sc.SetSheetCellAsDouble(h, input_val_col, base_row_val + from_high_row, from_high);

	if (from_low > 0)

		sc.SetSheetCellAsDouble(h, input_val_col, base_row_val + from_low_row, from_low);

}