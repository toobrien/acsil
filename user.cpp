#include "sierrachart.h"
#include <string>
#include <stdlib.h>
#include <map>

SCDLLName("user")

SCSFExport scsf_order_flow(SCStudyInterfaceRef sc) {

	#define MAX_SYMBOL_ROWS 1000

	#define input_head_col  0
	#define input_val_col	1
	#define stat_head_col   2
	#define stat_val_col	3

	#define symbol_row			0
	#define trades_row			1
	#define liq_lvls_row		2
	#define min_rotation_row	3

	#define liquidity_balance_row	0
	#define delta_row				1
	#define imbalance_row			2
	#define ask_tick_avg_row		3
	#define bid_tick_avg_row		4
	#define range_density_row		5
	#define range_row				6
	#define net_ticks_row			7
	#define rotation_side_row		8
	#define rotation_start_row		9
	#define rotation_length_row		10
	#define volume_row				11
	#define sample_row				12

	#define base_row_key		0
	#define high_volume_key		1
	#define rotation_side_key	2
	#define rotation_high_key	3
	#define rotation_low_key	4

	// set defaults
	
	SCInputRef symbol_input	= sc.Input[0];
	SCInputRef file_input	= sc.Input[1];
	SCInputRef sheet_input	= sc.Input[2];

	int		&	base_row 			= sc.GetPersistentInt(base_row_key);
	int 	&	high_volume			= sc.GetPersistentInt(high_volume_key);
	int		&	rotation_side		= sc.GetPersistentInt(rotation_side_key);
	double 	&	rotation_high		= sc.GetPersistentDouble(rotation_high_key);
	double 	&	rotation_low		= sc.GetPersistentDouble(rotation_low_key);

	if (sc.SetDefaults) {

		sc.GraphName 			= "order_flow";
		sc.AutoLoop 			= 0;
		sc.UsesMarketDepthData 	= 1;
		sc.HideStudy 			= 1;

		symbol_input.Name = "symbol";
		symbol_input.SetString("");

		file_input.Name = "file_name";
		file_input.SetString("");

		sheet_input.Name = "sheet_name";
		sheet_input.SetString("");
		
		base_row 			= -1;
		high_volume			= -1;
		rotation_side		= 0;
		rotation_high		= DBL_MIN;
		rotation_low		= DBL_MAX;

		return;

	}

	// initialize base row by scanning spreadsheet for input symbol

	const char * file_name	= file_input.GetString();
	const char * sheet_name = sheet_input.GetString();

	if (
		!std::strcmp(file_name, "")		||
		!std::strcmp(sheet_name, "")
	)

		// user has not initialized the spreadsheet inputs

		return;

	void * h = sc.GetSpreadsheetSheetHandleByName(file_name, sheet_name, false);

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

	// initialize remaining inputs from spreadsheet

	double d_trades			= -1.0;
	double d_liq_lvls 		= -1.0;
	double d_min_rotation	= -1.0;

	sc.GetSheetCellAsDouble(h, input_val_col, base_row + trades_row, d_trades);
	sc.GetSheetCellAsDouble(h, input_val_col, base_row + liq_lvls_row, d_liq_lvls);
	sc.GetSheetCellAsDouble(h, input_val_col, base_row + min_rotation_row, d_min_rotation);

	int trades			= static_cast<int>(d_trades);
	int liq_levels  	= static_cast<int>(d_liq_lvls);
	int min_rotation 	= static_cast<int>(d_min_rotation); 
	
	c_SCTimeAndSalesArray tas;

	if (trades > 0)

		sc.GetTimeAndSales(tas);

	// compute stats

	int len_tas	= tas.Size();

	double first_price		= 0.0;
	double last_price		= 0.0;
	double prev_ask			= 0.0;
	double prev_bid			= 0.0;
	double ask_ticks		= 0.0;
	double bid_ticks		= 0.0;
	double at_bid_total		= 0.0;
	double at_ask_total		= 0.0;
	double imbalance		= 0.0;
	double net_ticks		= 0.0;
	double ask_tick_avg		= 0.0;
	double bid_tick_avg		= 0.0;
	double high_tick		= DBL_MIN;
	double low_tick			= DBL_MAX;
	double range			= 0.0;
	double range_density	= 0.0;
	double sample			= 0.0;

	double delta 	= -1.0;
	double volume	= 0.0;

	double rotation_length = -1;

	bool init_ts = false;

	if (len_tas > 0) {

		int trade_count = 0;
		int start = max(len_tas - trades, 0);

		for (int i = start; i < len_tas; i++) {

			// init

			s_TimeAndSales r = tas[i];
			r *= sc.RealTimePriceMultiplier;

			if (!init_ts && r.Type != SC_TS_BIDASKVALUES) {

				// first trade record, initialize everything

				first_price = r.Price;
				prev_bid 	= r.Price;
				prev_ask	= r.Price;

				init_ts = true;

			}

			// range, absorbtion, density, etc.

			switch(r.Type) {

				case SC_TS_BID:

					at_bid_total	+=	r.Volume;
					last_price		=	r.Price;
					trade_count		+=	1;
					low_tick		=	min(r.Price, low_tick);

					if (r.Price < prev_bid)

						bid_ticks += ((prev_bid - r.Price) / sc.TickSize);

					prev_bid = r.Price;					

					// sc.AddMessageToLog(("at_bid: " + std::to_string(at_bid)).c_str(), 1);

					break;

				case SC_TS_ASK:

					at_ask_total	+= 	r.Volume;
					last_price		=	r.Price;
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

			// rotations

			if (r.Type != SC_TS_BIDASKVALUES) {

				rotation_high 	= max(rotation_high, r.Price);
				rotation_low	= min(rotation_low, r.Price);

				double from_rotation_high 	= (rotation_high - r.Price) / sc.TickSize;
				double from_rotation_low	= (r.Price - rotation_low) / sc.TickSize;

				if (from_rotation_high >= min_rotation) {
					
					if (rotation_side > -1) {

						rotation_side 		= -1;
						rotation_low 		= r.Price;
						from_rotation_low 	= 0;			// skip subsequent if block

					}
					
					rotation_length = max(from_rotation_high, rotation_length);
				
				}

				if (from_rotation_low >= min_rotation) {

					if (rotation_side < 1) {

						rotation_side	= 1;
						rotation_high	= r.Price;
 
					}

					rotation_length = max(from_rotation_low, rotation_length);

				}

			}

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

	// compute liquidity balance

	int 	total_bids 			= 0;
	int 	total_asks 			= 0;
	float	liquidity_balance 	= -1.0;

	s_MarketDepthEntry e;

	for (int i = 0; i < liq_levels; i++) {

		sc.GetBidMarketDepthEntryAtLevel(e, i);
		total_bids += e.Quantity;

		sc.GetAskMarketDepthEntryAtLevel(e, i);
		total_asks += e.Quantity;

	}

	liquidity_balance = static_cast<float>(total_asks) / total_bids;

	// clear spreadsheet

	const int 	lo = base_row + 1;
	int			hi = base_row + sample_row;
	SCString 	clr = "";
	SCString 	fmt;

	sc.SetSheetCellAsString(h, stat_val_col, base_row + liquidity_balance_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + delta_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + imbalance_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + ask_tick_avg_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + bid_tick_avg_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + range_density_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + range_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + net_ticks_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + rotation_side_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + rotation_start_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + rotation_length_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + volume_row, clr);
	sc.SetSheetCellAsString(h, stat_val_col, base_row + sample_row, clr);

	// fill spreadsheet

	sc.SetSheetCellAsString(h, stat_val_col, base_row + liquidity_balance_row, fmt.Format("%.2f", liquidity_balance));
	sc.SetSheetCellAsString(h, stat_val_col, base_row + delta_row, fmt.Format("%.2f", delta));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + imbalance_row, imbalance);
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + ask_tick_avg_row, static_cast<int>(ask_tick_avg));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + bid_tick_avg_row, static_cast<int>(bid_tick_avg));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + range_density_row, static_cast<int>(range_density));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + range_row, static_cast<int>(range));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + net_ticks_row, static_cast<int>(net_ticks));
	sc.SetSheetCellAsString(h, stat_val_col, base_row + rotation_side_row, rotation_side == 1 ? "up" : rotation_side == -1 ? "dn" : "");
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + rotation_start_row, rotation_side == 1 ? rotation_low : rotation_side == -1 ? rotation_high : -1);
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + rotation_length_row, static_cast<int>(rotation_length));
	sc.SetSheetCellAsString(h, stat_val_col, base_row + volume_row, fmt.Format("%.2f", volume));
	sc.SetSheetCellAsDouble(h, stat_val_col, base_row + sample_row, sample);

}


/*

INCOMPLETE

SCSFExport scsf_large_orders(SCStudyInterfaceRef sc) {

	/*

	#define bid_lvls_col	0
	#define bid_qtys_col   	1
	#define bid_tics_col    2
	#define ask_lvls_col 	3
	#define ask_qtys_col   	4
	#define ask_tics_col	5

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

}

// INCOMPLETE

SCSFExport scsf_bond_tick(SCStudyInterfaceRef sc) {
	
	#define zb_idx	0
	#define	zn_idx	1
	#define zf_idx	2
	#define zt_idx  3

	SCInputRef sheet_name_input		= sc.Input[1];
	SCInputRef base_symbol_input	= sc.Input[2];	// ZN, ZF, etc...
	SCInputRef month_year_input		= sc.Input[0];	// MYY

	if (sc.SetDefaults) {

		sc.GraphName 			= "tick_ranges";
		sc.AutoLoop 			= 0;
		// sc.UsesMarketDepthData 	= 1;
		sc.HideStudy 			= 1;

		sheet_name_input.Name = "sheet_name";
		sheet_name_input.SetString("");

		base_symbol_input.Name = "base_symbol";
		base_symbol_input.SetString("");

		month_year_input.Name = "month_year";
		month_year_input.SetString("");

		return;
		
	}

	const char * 	sheet_name		= sheet_name_input.GetString();
	const char * 	base_symbol		= base_symbol_input.GetString();
	const char * 	month_year		= month_year_input.GetString();
	SCString 		clr				= "";

	if (
		!clr.Compare(sheet_name) 	||
		!clr.Compare(base_symbol)	||
		!clr.Compare(month_year)
	)

		// user has not initialized study

		return;

	std::map<double, double>* tick_ranges = reinterpret_cast<std::map<double, double>*>(sc.GetPersistentPointer(0));

	if (sc.LastCallToFunction) {

		if (tick_ranges != NULL) {

			delete tick_ranges;
			sc.SetPersistentPointer(0, NULL);

		}

		return;

	}

	if (tick_ranges == NULL)

		tick_ranges = new std::map<double, double>();

	void * h = sc.GetSpreadsheetSheetHandleByName("bonds", sheet_name, false);

	SCString fmt;

	SCString zb_sym = fmt.Format("ZB%S_FUT_CME", month_year);
	SCString zn_sym = fmt.Format("ZN%S_FUT_CME", month_year);
	SCString zf_sym = fmt.Format("ZF%S_FUT_CME", month_year);
	SCString zt_sym = fmt.Format("ZT%S_FUT_CME", month_year);

	c_SCTimeAndSalesArray zb_tas;
	c_SCTimeAndSalesArray zn_tas;
	c_SCTimeAndSalesArray zf_tas;
	c_SCTimeAndSalesArray zt_tas;

	sc.GetTimeAndSalesForSymbol(zb_sym, zb_tas);
	sc.GetTimeAndSalesForSymbol(zn_sym, zn_tas);
	sc.GetTimeAndSalesForSymbol(zf_sym, zf_tas);
	sc.GetTimeAndSalesForSymbol(zt_sym, zt_tas);

}

*/


SCSFExport scsf_tpo_to_spreadsheet(SCStudyInterfaceRef sc) {

	#define NUM_MEMBERS 51

	#define	StartDateTime_row						0
    #define NumberOfTrades_row						1
    #define Volume_row								2
    #define BidVolume_row							3
    #define AskVolume_row							4
    #define TotalTPOCount_row						5
    #define OpenPrice_row							6
    #define HighestPrice_row						7
    #define LowestPrice_row							8
	#define LastPrice_row							9
    #define TPOMidpointPrice_row					10
    #define TPOMean_row								11
    #define TPOStdDev_row							12
    #define TPOErrorOfMean_row						13
    #define TPOPOCPrice_row							14
    #define TPOValueAreaHigh_row					15
    #define TPOValueAreaLow_row						16
    #define TPOCountAbovePOC_row					17
    #define TPOCountBelowPOC_row					18
    #define VolumeMidpointPrice_row					19
    #define VolumePOCPrice_row						20
    #define VolumeValueAreaHigh_row					21
    #define VolumeValueAreaLow						22
    #define VolumeAbovePOC_row						23
    #define VolumeBelowPOC_row						24
    #define POCAboveBelowVolumeImbalancePercent_row	25
    #define VolumeAboveLastPrice_row				26
    #define VolumeBelowLastPrice_row				27
    #define BidVolumeAbovePOC_row					28
    #define BidVolumeBelowPOC_row					29
    #define AskVolumeAbovePOC_row					30
    #define AskVolumeBelowPOC_row					31
    #define VolumeTimesPriceInTicks_row				32
	#define TradesTimesPriceInTicks_row				33
    #define TradesTimesPriceSquaredInTicks_row		34
    #define IBRHighPrice_row						35
    #define IBRLowPrice_row							36
    #define OpeningRangeHighPrice_row				37
    #define OpeningRangeLowPrice_row				38
    #define VolumeWeightedAveragePrice_row			39
    #define MaxTPOBlocksCount_row					40
    #define TPOCountMaxDigits_row					41
    #define DisplayIndependentColumns_row			42
    #define EveningSession_row						43
    #define AverageSubPeriodRange_row				44
    #define RotationFactor_row						45
    #define VolumeAboveTPOPOC_row					46
    #define VolumeBelowTPOPOC_row					47
    #define EndDateTime_row							48
    #define BeginIndex_row							49
    #define EndIndex_row							50

	SCInputRef file_name_input 		= sc.Input[0];  // spreadsheet file name
	SCInputRef sheet_name_input		= sc.Input[1];  // specific sheet
	SCInputRef num_profiles_input	= sc.Input[2];	// max number of profiles to use
	SCInputRef study_id_input		= sc.Input[3];  // TPO study id

	int & initialized = sc.GetPersistentInt(0);

	if (sc.SetDefaults) {

		sc.GraphName 			= "tpo_to_spreadsheet";
		sc.AutoLoop 			= 0;
		sc.HideStudy 			= 1;

		file_name_input.Name = "file_name";
		file_name_input.SetString("");

		sheet_name_input.Name = "sheet_name";
		sheet_name_input.SetString("");

		num_profiles_input.Name = "num_profiles";
		num_profiles_input.SetInt(-1);

		study_id_input.Name = "study_id";
		study_id_input.SetInt(-1);

		initialized = 0;

		return;
		
	}

	const char * 	file_name 		= file_name_input.GetString();
	const char * 	sheet_name		= sheet_name_input.GetString();
	SCString		clr				= "";
	int 			num_profiles 	= num_profiles_input.GetInt();
	int				study_id		= study_id_input.GetInt();

	if (
		!clr.Compare(file_name) 	||
		!clr.Compare(sheet_name)	||
		num_profiles < 0			||
		study_id     < 0
	)

		// user has not initialized study

		return;

	void * h = sc.GetSpreadsheetSheetHandleByName(file_name, sheet_name, false);

	// set row headers

	if (!initialized) {

		sc.SetSheetCellAsString(h, 0, StartDateTime_row, "StartDateTime");
		sc.SetSheetCellAsString(h, 0, NumberOfTrades_row, "NumberOfTrades");
		sc.SetSheetCellAsString(h, 0, Volume_row, "Volume");
		sc.SetSheetCellAsString(h, 0, BidVolume_row, "BidVolume");
		sc.SetSheetCellAsString(h, 0, AskVolume_row, "AskVolume");
		sc.SetSheetCellAsString(h, 0, TotalTPOCount_row, "TotalTPOCount");
		sc.SetSheetCellAsString(h, 0, OpenPrice_row, "OpenPrice");
		sc.SetSheetCellAsString(h, 0, HighestPrice_row, "Highest_price");
		sc.SetSheetCellAsString(h, 0, LowestPrice_row, "LowestPrice");
		sc.SetSheetCellAsString(h, 0, LastPrice_row, "LastPrice");
		sc.SetSheetCellAsString(h, 0, TPOMidpointPrice_row, "TPOMidpointPrice");
		sc.SetSheetCellAsString(h, 0, TPOMean_row, "TPOMean");
		sc.SetSheetCellAsString(h, 0, TPOStdDev_row, "TPOStdDev");
		sc.SetSheetCellAsString(h, 0, TPOErrorOfMean_row, "TPOErrorOfMean");
		sc.SetSheetCellAsString(h, 0, TPOMean_row, "TPOMean");
		sc.SetSheetCellAsString(h, 0, TPOPOCPrice_row, "TPOPOCPrice");
		sc.SetSheetCellAsString(h, 0, TPOValueAreaHigh_row, "TPOValueAreaHigh");
		sc.SetSheetCellAsString(h, 0, TPOValueAreaLow_row, "TPOValueAreaLow");
		sc.SetSheetCellAsString(h, 0, TPOCountAbovePOC_row, "TPOCountAbovePOC");
		sc.SetSheetCellAsString(h, 0, TPOCountBelowPOC_row, "TPOCountBelowPOC");
		sc.SetSheetCellAsString(h, 0, VolumeMidpointPrice_row, "VolumeMidpointPrice");
		sc.SetSheetCellAsString(h, 0, VolumePOCPrice_row, "VolumePOCPrice");
		sc.SetSheetCellAsString(h, 0, VolumeValueAreaHigh_row, "VolumeValueAreaHigh");
		sc.SetSheetCellAsString(h, 0, VolumeValueAreaLow, "VolumeValueAreaLow");
		sc.SetSheetCellAsString(h, 0, VolumeAbovePOC_row, "VolumeAbovePOC");
		sc.SetSheetCellAsString(h, 0, VolumeBelowPOC_row, "VolumeBelowPOC");
		sc.SetSheetCellAsString(h, 0, POCAboveBelowVolumeImbalancePercent_row, "POCAboveBelowVolumeImbalancePercent");
		sc.SetSheetCellAsString(h, 0, VolumeAboveLastPrice_row, "VolumeAboveLastPrice");
		sc.SetSheetCellAsString(h, 0, VolumeBelowLastPrice_row, "VolumeBelowLastPrice");
		sc.SetSheetCellAsString(h, 0, BidVolumeAbovePOC_row, "BidVolumeAbovePOC");
		sc.SetSheetCellAsString(h, 0, BidVolumeBelowPOC_row, "BidVolumeBelowPOC");
		sc.SetSheetCellAsString(h, 0, AskVolumeAbovePOC_row, "AskVolumeAbovePOC");
		sc.SetSheetCellAsString(h, 0, AskVolumeBelowPOC_row, "AskVolumeBelowPOC");
		sc.SetSheetCellAsString(h, 0, VolumeTimesPriceInTicks_row, "VolumeTimesPriceInTicks");
		sc.SetSheetCellAsString(h, 0, TradesTimesPriceInTicks_row, "TradesTimesPriceInTicks");
		sc.SetSheetCellAsString(h, 0, TradesTimesPriceSquaredInTicks_row, "TradesTimesPriceSquaredInTicks");
		sc.SetSheetCellAsString(h, 0, IBRHighPrice_row, "IBRHighPrice");
		sc.SetSheetCellAsString(h, 0, IBRLowPrice_row, "IBRLowPrice");
		sc.SetSheetCellAsString(h, 0, OpeningRangeHighPrice_row, "OpeningRangeHighPrice");
		sc.SetSheetCellAsString(h, 0, OpeningRangeLowPrice_row, "OpeningRangeLowPrice");
		sc.SetSheetCellAsString(h, 0, VolumeWeightedAveragePrice_row, "VolumeWeightedAveragePrice");
		sc.SetSheetCellAsString(h, 0, MaxTPOBlocksCount_row, "MaxTPOBlocksCount");
		sc.SetSheetCellAsString(h, 0, TPOCountMaxDigits_row, "TPOCountMaxDigits");
		sc.SetSheetCellAsString(h, 0, DisplayIndependentColumns_row, "DisplayIndependentColumns");
		sc.SetSheetCellAsString(h, 0, EveningSession_row, "EveningSession");
		sc.SetSheetCellAsString(h, 0, AverageSubPeriodRange_row, "AverageSubPeriodRange");
		sc.SetSheetCellAsString(h, 0, RotationFactor_row, "RotationFactor");
		sc.SetSheetCellAsString(h, 0, VolumeAboveTPOPOC_row, "VolumeAboveTPOPOC");
		sc.SetSheetCellAsString(h, 0, VolumeBelowTPOPOC_row, "VolumeBelowTPOPOC");
		sc.SetSheetCellAsString(h, 0, EndDateTime_row, "EndDateTime");
		sc.SetSheetCellAsString(h, 0, BeginIndex_row, "BeginIndex");
		sc.SetSheetCellAsString(h, 0, EndIndex_row, "EndIndex");
	
	}
    
	// copy TPO values into spreadsheet

	int j = 1; // column

	for (int i = 0; i < num_profiles; i++) {

		n_ACSIL::s_StudyProfileInformation p;

		int res = sc.GetStudyProfileInformation(study_id, i, p);
		
		if (res) {
			
			// sc.SetSheetCellAsString(h, j, StartDateTime_row, p.m_StartDateTime);
			sc.SetSheetCellAsDouble(h, j, NumberOfTrades_row, p.m_NumberOfTrades);
			sc.SetSheetCellAsDouble(h, j, Volume_row, p.m_Volume);
			sc.SetSheetCellAsDouble(h, j, BidVolume_row, p.m_BidVolume);
			sc.SetSheetCellAsDouble(h, j, AskVolume_row, p.m_AskVolume);
			sc.SetSheetCellAsDouble(h, j, TotalTPOCount_row, p.m_TotalTPOCount);
			sc.SetSheetCellAsDouble(h, j, OpenPrice_row, p.m_OpenPrice);
			sc.SetSheetCellAsDouble(h, j, HighestPrice_row, p.m_HighestPrice);
			sc.SetSheetCellAsDouble(h, j, LowestPrice_row, p.m_LowestPrice);
			sc.SetSheetCellAsDouble(h, j, LastPrice_row, p.m_LastPrice);
			sc.SetSheetCellAsDouble(h, j, TPOMidpointPrice_row, p.m_TPOMidpointPrice);
			sc.SetSheetCellAsDouble(h, j, TPOMean_row, p.m_TPOMean);
			sc.SetSheetCellAsDouble(h, j, TPOStdDev_row, p.m_TPOStdDev);
			sc.SetSheetCellAsDouble(h, j, TPOErrorOfMean_row, p.m_TPOErrorOfMean);
			sc.SetSheetCellAsDouble(h, j, TPOMean_row, p.m_TPOMean);
			sc.SetSheetCellAsDouble(h, j, TPOPOCPrice_row, p.m_TPOPOCPrice);
			sc.SetSheetCellAsDouble(h, j, TPOValueAreaHigh_row, p.m_TPOValueAreaHigh);
			sc.SetSheetCellAsDouble(h, j, TPOValueAreaLow_row, p.m_TPOValueAreaLow);
			sc.SetSheetCellAsDouble(h, j, TPOCountAbovePOC_row, p.m_TPOCountAbovePOC);
			sc.SetSheetCellAsDouble(h, j, TPOCountBelowPOC_row, p.m_TPOCountBelowPOC);
			sc.SetSheetCellAsDouble(h, j, VolumeMidpointPrice_row, p.m_VolumeMidpointPrice);
			sc.SetSheetCellAsDouble(h, j, VolumePOCPrice_row, p.m_VolumePOCPrice);
			sc.SetSheetCellAsDouble(h, j, VolumeValueAreaHigh_row, p.m_VolumeValueAreaHigh);
			sc.SetSheetCellAsDouble(h, j, VolumeValueAreaLow, p.m_VolumeValueAreaLow);
			sc.SetSheetCellAsDouble(h, j, VolumeAbovePOC_row, p.m_VolumeAbovePOC);
			sc.SetSheetCellAsDouble(h, j, VolumeBelowPOC_row, p.m_VolumeBelowPOC);
			sc.SetSheetCellAsDouble(h, j, POCAboveBelowVolumeImbalancePercent_row, p.m_POCAboveBelowVolumeImbalancePercent);
			sc.SetSheetCellAsDouble(h, j, VolumeAboveLastPrice_row, p.m_VolumeAboveLastPrice);
			sc.SetSheetCellAsDouble(h, j, VolumeBelowLastPrice_row, p.m_VolumeBelowLastPrice);
			sc.SetSheetCellAsDouble(h, j, BidVolumeAbovePOC_row, p.m_BidVolumeAbovePOC);
			sc.SetSheetCellAsDouble(h, j, BidVolumeBelowPOC_row, p.m_BidVolumeBelowPOC);
			sc.SetSheetCellAsDouble(h, j, AskVolumeAbovePOC_row, p.m_AskVolumeAbovePOC);
			sc.SetSheetCellAsDouble(h, j, AskVolumeBelowPOC_row, p.m_AskVolumeBelowPOC);
			sc.SetSheetCellAsDouble(h, j, VolumeTimesPriceInTicks_row, p.m_VolumeTimesPriceInTicks);
			sc.SetSheetCellAsDouble(h, j, TradesTimesPriceInTicks_row, p.m_TradesTimesPriceInTicks);
			sc.SetSheetCellAsDouble(h, j, TradesTimesPriceSquaredInTicks_row, p.m_TradesTimesPriceSquaredInTicks);
			sc.SetSheetCellAsDouble(h, j, IBRHighPrice_row, p.m_IBRHighPrice);
			sc.SetSheetCellAsDouble(h, j, IBRLowPrice_row, p.m_IBRLowPrice);
			sc.SetSheetCellAsDouble(h, j, OpeningRangeHighPrice_row, p.m_OpeningRangeHighPrice);
			sc.SetSheetCellAsDouble(h, j, OpeningRangeLowPrice_row, p.m_OpeningRangeLowPrice);
			sc.SetSheetCellAsDouble(h, j, VolumeWeightedAveragePrice_row, p.m_VolumeWeightedAveragePrice);
			sc.SetSheetCellAsDouble(h, j, MaxTPOBlocksCount_row, p.m_MaxTPOBlocksCount);
			sc.SetSheetCellAsDouble(h, j, TPOCountMaxDigits_row, p.m_TPOCountMaxDigits);
			// sc.SetSheetCellAsDouble(h, j, DisplayIndependentColumns_row, p.m_DisplayIndependentColumns);
			sc.SetSheetCellAsDouble(h, j, EveningSession_row, p.m_EveningSession);
			sc.SetSheetCellAsDouble(h, j, AverageSubPeriodRange_row, p.m_AverageSubPeriodRange);
			sc.SetSheetCellAsDouble(h, j, RotationFactor_row, p.m_RotationFactor);
			sc.SetSheetCellAsDouble(h, j, VolumeAboveTPOPOC_row, p.m_VolumeAboveTPOPOC);
			sc.SetSheetCellAsDouble(h, j, VolumeBelowTPOPOC_row, p.m_VolumeBelowTPOPOC);
			// sc.SetSheetCellAsString(h, j, EndDateTime_row, p.m_EndDateTime);
			sc.SetSheetCellAsDouble(h, j, BeginIndex_row, p.m_BeginIndex);
			sc.SetSheetCellAsDouble(h, j, EndIndex_row, p.m_EndIndex);

			j++;
	
		}

		if (initialized)

			// only update the latest profile after the first run.
			// need to restart the study after a new session.

			break;
	
	}

}