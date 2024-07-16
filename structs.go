package main

import (
	"math"
)

type RawData struct {
	FeatID       string  `json:"Feat ID"`
	OldComp      string  `json:"Old Comp"`
	Survey       int     `json:"Survey"`
	PlotNo       string  `json:"Plot No"`
	MainStem     float64 `json:"Main Stem"`
	SecondStem   float64 `json:"Second Stem"`
	ThirdStem    float64 `json:"Third Stem"`
	FourthStem   float64 `json:"Fourth Stem"`
	Height       float64 `json:"Height"`
	SecondHeight float64 `json:"Second Height"`
	ThirdHeight  float64 `json:"Third Height"`
	FourthHeight float64 `json:"Fourth Height"`
	// DeadStem   float64 `json:"Dead Stem"`
	Remark string `json:"Remark"`
}

type ConvertedData struct {
	FeatID         string
	OldComp        string
	Survey         int
	PlotNo         string
	DBH            float64
	MainStem       int
	SecondStem     int
	FallenStem     int
	SickStem       int
	DR             int
	DrFreq         int
	DeadSt         int
	HT1            float64
	PlanHarvesting string
	Remark         string
}

type AllRawData struct {
	AllData []RawData
}

func (r *AllRawData) FormatRow() *AllRawData {
	for i, v := range r.AllData {
		if v.SecondStem > 1 {
			newRow := RawData{
				FeatID:    v.FeatID,
				OldComp:   v.OldComp,
				Survey:    v.Survey,
				PlotNo:    v.PlotNo,
				MainStem:  v.SecondStem,
				ThirdStem: 1,
				Height:    v.SecondHeight,
			}
			r.AllData = append(r.AllData, newRow)
			r.AllData[i].SecondStem = 0
		}
		if v.ThirdStem > 1 {
			newRow := RawData{
				FeatID:    v.FeatID,
				OldComp:   v.OldComp,
				Survey:    v.Survey,
				PlotNo:    v.PlotNo,
				MainStem:  v.ThirdStem,
				ThirdStem: 1,
				Height:    v.ThirdHeight,
			}
			r.AllData = append(r.AllData, newRow)
			r.AllData[i].ThirdStem = 0
		}
		if v.FourthStem > 1 {
			newRow := RawData{
				FeatID:    v.FeatID,
				OldComp:   v.OldComp,
				Survey:    v.Survey,
				PlotNo:    v.PlotNo,
				MainStem:  v.FourthStem,
				ThirdStem: 1,
				Height:    v.FourthHeight,
			}
			r.AllData = append(r.AllData, newRow)
			r.AllData[i].FourthStem = 0
		}
	}
	return r
}

func (r *AllRawData) GroupBy() map[float64][]RawData {
	groupedDBH := make(map[float64][]RawData, 0)
	for _, v := range r.AllData {
		groupedDBH[v.MainStem] = append(groupedDBH[v.MainStem], v)
	}
	return groupedDBH
}

func roundFloat(val float64, precision uint) float64 {
	ratio := math.Pow(10, float64(precision))
	return math.Round(val*ratio) / ratio
}
