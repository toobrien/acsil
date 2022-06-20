#include "sierrachart.h"
#include <string>
#include <stdlib.h>

// For reference, refer to this page:
// https://www.sierrachart.com/index.php?page=doc/AdvancedCustomStudyInterfaceAndLanguage.php

SCDLLName("depth_studies")

SCSFExport scsf_large_orders(SCStudyInterfaceRef sc) {

	#define MAX_SYMBOL_ROWS 1000

	#define input_head_col  0
	#define input_val_col	1
	#define spacer_col_a	2
	#define stat_head_col   3
	#define stat_val_col	4
	#define spacer_col_b	5
	#define bid_lvls_col	6
	#define bid_qtys_col   	7
	#define bid_tics_col    8
	#define spacer_col_c    9
	#define ask_lvls_col 	10
	#define ask_qtys_col   	11
	#define ask_tics_col	12
	
	#define symbol_row			0
	#define max_outputs_row		1
	#define threshold_row		2
	#define alert_distance_row	3	// not used
	#define num_trades_row		4
	#define num_liq_lvls_row	5

	#define from_high_row			0
	#define from_low_row			1
	#define liquidity_balance_row	2
	#define delta_row				3
	#define imbalance_row			4
	#define ask_tick_avg_row		5
	#define bid_tick_avg_row		6
	#define range_density_row		7
	#define range_row				8
	#define net_ticks_row			9
	#define volume_row				10
	#define sample_row				11

	#define base_row_key			0
	#define high_volume_key			1

	// set defaults
	
	SCInputRef symbol_input	= sc.Input[0];
	SCInputRef sheet_input	= sc.Input[1];

	int	&	base_row 		= sc.GetPersistentInt(base_row_key);
	int &	high_volume		= sc.GetPersistentInt(high_volume_key);

	if (sc.SetDefaults) {

		sc.GraphName 			= "large_orders";
		sc.AutoLoop 			= 0;
		sc.UsesMarketDepthData 	= 1;
		sc.HideStudy 			= 1;

		symbol_input.Name = "symbol";
		symbol_input.SetString("");

		sheet_input.Name = "sheet name";
		sheet_input.SetString("");
		
		base_row 		= -1;
		high_volume		= -1;

		return;
		
	}

	// initialize base row by scanning spreadsheet for input symbol

	const char * sheet_name = sheet_input.GetString();

	if (!std::strcmp(sheet_name, ""))

		// user has not set sheet name

		return;

	void * h = sc.GetSpreadsheetSheetHandleByName("depth_sheet", sheet_name, false);

	if (base_row < 0) {
		
		const char * 	symbol = symbol_input.GetString(); 
		SCString 		s;

		for (int i = 0; i < MAX_SYMBOL_ROWS; i++) {

			sc.GetSheetCellAsString(h, input_val_col, i, s);

			if (!s.Compare(symbol)) {

				base_row = i;
				break;

			}
		}

		if (base_row < 0)

			// user has not selected a symbol or spreadsheet is too large

			return;

	}

	// START	initialize remaining inputs from spreadsheet

	double d_max_outputs 	= -1.0;
	double d_threshold 		= -1.0;
	double d_alert_distance = -1.0;
	double d_num_trades		= -1.0;
	double d_num_liq_lvls 	= -1.0;
	
	sc.GetSheetCellAsDouble(h, input_val_col, base_row + max_outputs_row, d_max_outputs);
	sc.GetSheetCellAsDouble(h, input_val_col, base_row + threshold_row, d_threshold);
	sc.GetSheetCellAsDouble(h, input_val_col, base_row + alert_distance_row, d_alert_distance);
	sc.GetSheetCellAsDouble(h, input_val_col, base_row + num_trades_row, d_num_trades);
	sc.GetSheetCellAsDouble(h, input_val_col, base_row + num_liq_lvls_row, d_num_liq_lvls);

	int max_outputs 	= static_cast<int>(d_max_outputs);
	int threshold		= static_cast<int>(d_threshold);
	int alert_distance	= static_cast<int>(d_alert_distance);
	int num_trades		= static_cast<int>(d_num_trades);
	int num_liq_levels  = static_cast<int>(d_num_liq_lvls);

	c_SCTimeAndSalesArray tas;

	if (num_trades > 0)

		sc.GetTimeAndSales(tas);

	// END		initialize remaining inputs from spreadsheet

	// START 	time and sales: delta, volume

	int 	len_tas	= tas.Size();

	double	first_price		= DBL_MIN;
	double	last_price		= DBL_MIN;
	double  prev_ask		= DBL_MAX;
	double  prev_bid		= DBL_MIN;
	double  ask_ticks		= 0.0;
	double 	bid_ticks		= 0.0;
	double 	at_bid_total	= 0.0;
	double 	at_ask_total	= 0.0;
	double	imbalance		= 0.0;
	double	net_ticks		= 0.0;
	double	ask_tick_avg	= 0.0;
	double  bid_tick_avg	= 0.0;
	double  high_tick		= DBL_MIN;
	double  low_tick		= DBL_MAX;
	double	range			= 0.0;
	double  range_density	= 0.0;
	double	sample			= 0.0;

	double 	delta 	= -1.0;
	double 	volume	= 0.0;

	if (len_tas > 0) {

		// record volumes, last price

		int i 			= len_tas - 1;
		int trade_count = 0;

		while (trade_count < num_trades && i >= 0) {

			s_TimeAndSales r = tas[i];
			r *= sc.RealTimePriceMultiplier;

			switch(r.Type) {

				case SC_TS_BID:

					at_bid_total	+=	r.Volume;
					first_price		=	r.Price;
					trade_count		+=	1;
					low_tick		=	min(r.Price, low_tick);

					//sc.AddMessageToLog(("cur_bid: " + std::to_string(r.Price)).c_str(), 1);

					if (r.Price < prev_bid) {

						bid_ticks += ((prev_bid - r.Price) / sc.TickSize);

						//sc.AddMessageToLog(("prev_bid: " + std::to_string(prev_bid)).c_str(), 1);
						//sc.AddMessageToLog(("ticks: " + std::to_string(((prev_bid - r.Price) / sc.TickSize))).c_str(), 1);

					}

					prev_bid = r.Price;					

					// sc.AddMessageToLog(("at_bid: " + std::to_string(at_bid)).c_str(), 1);

					break;

				case SC_TS_ASK:

					at_ask_total	+= 	r.Volume;
					first_price		=	r.Price;
					trade_count		+=	1;
					high_tick		=	max(r.Price, high_tick);

					if (r.Price > prev_ask)

						ask_ticks += ((r.Price - prev_ask) / sc.TickSize);

					prev_ask = r.Price;

					// sc.AddMessageToLog(("at_ask" + std::to_string(at_ask)).c_str(), 1);
					
					break;

				case SC_TS_BIDASKVALUES:

					break;

				default:

					break;
				
			}

			if (last_price == DBL_MIN && first_price != DBL_MIN)

				last_price = first_price;

			i--;

		}

		// compute normalized delta

		if (at_bid_total > 0.0)
		
			delta = at_ask_total / at_bid_total;

		volume = at_bid_total + at_ask_total;

		// compute normalized volume

		if (volume > high_volume)

			high_volume = volume;

		volume /= high_volume;

		// compute total ticks, net imbalance		

		net_ticks	= (last_price - first_price) / sc.TickSize;
		imbalance	= at_ask_total - at_bid_total;

		// compute average ticks

		if (ask_ticks > 0)

			ask_tick_avg = at_ask_total / ask_ticks;

		if (bid_ticks > 0)

			bid_tick_avg = at_bid_total / bid_ticks;

		// compute sample (lots in num_trades)

		sample = at_bid_total + at_ask_total;

		if (high_tick - low_tick > 0) {

			range			= (high_tick - low_tick) / sc.TickSize;
			range_density 	= (at_ask_total + at_bid_total) / range;

		}

	}

	// END		time and sales: delta, volume

	// START	large order levels, liquidity balance, from high/low

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

	int from_high 	= -1;
	int from_low 	= -1;

	int 	total_bids 			= 0;
	int 	total_asks 			= 0;
	float	liquidity_balance 	= -1.0;

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

		if (i < num_liq_levels)

			total_bids += e.Quantity;

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

		if (i < num_liq_levels)

			total_asks += e.Quantity;

		if (len_asks >= max_outputs)

			break;

	}

	liquidity_balance = static_cast<float>(total_asks) / total_bids;

	// END		large order levels, liquidity balance, from high/low

	// START	spreadsheet

	// clear spreadsheet

	const int 	lo = base_row + 1;
	int			hi = base_row + max_outputs;
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

	sc.SetSheetCellAsString(h, stat_val_col, base_row + from_high_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + from_low_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + liquidity_balance_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + net_ticks_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + imbalance_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + ask_tick_avg_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + bid_tick_avg_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + range_density_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + range_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + sample_row, clr);
	
	// put all bids and asks where qty > threshold into spreadsheet

	SCString fmt;

	hi = base_row + len_bids;
	int j = 0;

	for (int i = lo; i < hi; i++) {

		sc.SetSheetCellAsDouble(h, bid_lvls_col, i, bid_lvls[j]);
		sc.SetSheetCellAsDouble(h, bid_qtys_col, i, bid_qtys[j]);
		sc.SetSheetCellAsDouble(h, bid_tics_col, i, bid_tics[j++]);

	}

	hi = base_row + len_asks;
	j = 0;

	for (int i = lo; i < hi; i++) {

		sc.SetSheetCellAsDouble(h, ask_lvls_col, i, ask_lvls[j]);
		sc.SetSheetCellAsDouble(h, ask_qtys_col, i, ask_qtys[j]);
		sc.SetSheetCellAsDouble(h, ask_tics_col, i, ask_tics[j++]);

	}

	// set stats

	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + from_high_row, from_high);
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + from_low_row, from_low);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + liquidity_balance_row, fmt.Format("%.2f", liquidity_balance));
	sc.SetSheetCellAsString(h, stat_val_col, base_row + delta_row, fmt.Format("%.2f", delta));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + ask_tick_avg_row, static_cast<int>(ask_tick_avg));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + bid_tick_avg_row, static_cast<int>(bid_tick_avg));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + net_ticks_row, static_cast<int>(net_ticks));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + imbalance_row, imbalance);
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + range_density_row, static_cast<int>(range_density));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + range_row, static_cast<int>(range));
	sc.SetSheetCellAsString(h, stat_val_col, base_row + volume_row, fmt.Format("%.2f", volume));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + sample_row, sample);

	// END		spreadsheet

}