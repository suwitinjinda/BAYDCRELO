import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton, getIcon } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";

import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField';
import { Stack, IStackTokens } from '@fluentui/react';
// import { useLocalStorage } from "./useLocalStorage";
// const request = require('request')
const stackTokens = () => { };
/* global console, Excel, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [], start: 0, check: false
    };
  }

  componentDidMount() {
    this.setState({
      check: false
    });
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  //start dev
  active = (a) => {
    // console.log(a)
    if (a == 1)
      this.setState({ check: true });

    var interval = setInterval(() => {
      console.log('This will run every second!');
      // this.setState.check == true;
      checkstatus()
      // let sname = document.getElementById("datasheet").value;
      // console.log(sname)
    }, 5000);
    // return () => clearInterval(interval);
    const checkstatus = async () => {
      // console.log(document.getElementById("datasheet").value);
      let sname = document.getElementById("datasheet").value;
      let tname = document.getElementById("dashboard").value;
      let inprogenable = document.getElementById("inprogressenable").value;
      let inprogressaddress = document.getElementById("inprogress").value;
      let compenable = document.getElementById("completeenable").value;
      let completeaddress = document.getElementById("complete").value;
      let delayenable = document.getElementById("delayenable").value;
      let delayaddress = document.getElementById("delay").value;
      let icompenable = document.getElementById("incompleteenable").value;
      let Icompleteaddress = document.getElementById("incomplete").value;
      let cdelayenable = document.getElementById("completedelayenable").value;
      let Cdelayaddress = document.getElementById("completedelay").value;

      localStorage.setItem("dashboard", JSON.stringify({
        sname: sname,
        tname: tname, inprogressaddress: inprogressaddress, completeaddress: completeaddress, delayaddress: delayaddress,
        Icompleteaddress: Icompleteaddress
      }))
      //Group1
      let statusg1 = document.getElementById("datastatusg1").value;
      let taskidg1 = document.getElementById("datataskidg1").value;
      let durationg1 = document.getElementById("datadurationg1").value;
      let startg1 = document.getElementById("datastartg1").value;
      let endg1 = document.getElementById("dataendg1").value;
      let resourceg1 = document.getElementById("dataresourceg1").value;
      let groupg1 = document.getElementById("datagroupg1").value;
      let targetg1 = document.getElementById("targetrowg1").value;
      let dayg1 = document.getElementById("dayg1").value;

      localStorage.setItem("g1config", JSON.stringify({
        statusg1: statusg1,
        taskidg1: taskidg1, durationg1: durationg1, startg1: startg1, endg1: endg1,
        resourceg1: resourceg1, groupg1: groupg1, targetg1: targetg1, dayg1: dayg1
      }))

      //Group2
      let statusg2 = document.getElementById("datastatusg2").value;
      let taskidg2 = document.getElementById("datataskidg2").value;
      let durationg2 = document.getElementById("datadurationg2").value;
      let startg2 = document.getElementById("datastartg2").value;
      let endg2 = document.getElementById("dataendg2").value;
      let resourceg2 = document.getElementById("dataresourceg2").value;
      let groupg2 = document.getElementById("datagroupg2").value;
      let targetg2 = document.getElementById("targetrowg2").value;
      let dayg2 = document.getElementById("dayg2").value;
      let activeg2 = document.getElementById("activeg2").value;

      localStorage.setItem("g2config", JSON.stringify({
        statusg2: statusg2,
        taskidg2: taskidg2, durationg2: durationg2, startg2: startg2, endg2: endg2,
        resourceg2: resourceg2, groupg2: groupg2, targetg2: targetg2, dayg2: dayg2
      }))
      //Group3
      let statusg3 = document.getElementById("datastatusg3").value;
      let taskidg3 = document.getElementById("datataskidg3").value;
      let durationg3 = document.getElementById("datadurationg3").value;
      let startg3 = document.getElementById("datastartg3").value;
      let endg3 = document.getElementById("dataendg3").value;
      let resourceg3 = document.getElementById("dataresourceg3").value;
      let groupg3 = document.getElementById("datagroupg3").value;
      let targetg3 = document.getElementById("targetrowg3").value;
      let dayg3 = document.getElementById("dayg3").value;
      let activeg3 = document.getElementById("activeg3").value;

      localStorage.setItem("g3config", JSON.stringify({
        statusg3: statusg3,
        taskidg3: taskidg3, durationg3: durationg3, startg3: startg3, endg3: endg3,
        resourceg3: resourceg3, groupg3: groupg3, targetg3: targetg3, dayg3: dayg3
      }))
      //Group4
      let statusg4 = document.getElementById("datastatusg4").value;
      let taskidg4 = document.getElementById("datataskidg4").value;
      let durationg4 = document.getElementById("datadurationg4").value;
      let startg4 = document.getElementById("datastartg4").value;
      let endg4 = document.getElementById("dataendg4").value;
      let resourceg4 = document.getElementById("dataresourceg4").value;
      let groupg4 = document.getElementById("datagroupg4").value;
      let targetg4 = document.getElementById("targetrowg4").value;
      let dayg4 = document.getElementById("dayg4").value;
      let activeg4 = document.getElementById("activeg4").value;

      localStorage.setItem("g4config", JSON.stringify({
        statusg4: statusg4,
        taskidg4: taskidg4, durationg4: durationg4, startg4: startg4, endg4: endg4,
        resourceg4: resourceg4, groupg4: groupg4, targetg4: targetg4, dayg4: dayg4
      }))
      //Group5
      let statusg5 = document.getElementById("datastatusg5").value;
      let taskidg5 = document.getElementById("datataskidg5").value;
      let durationg5 = document.getElementById("datadurationg5").value;
      let startg5 = document.getElementById("datastartg5").value;
      let endg5 = document.getElementById("dataendg5").value;
      let resourceg5 = document.getElementById("dataresourceg5").value;
      let groupg5 = document.getElementById("datagroupg5").value;
      let targetg5 = document.getElementById("targetrowg5").value;
      let dayg5 = document.getElementById("dayg5").value;
      let activeg5 = document.getElementById("activeg5").value;

      localStorage.setItem("g5config", JSON.stringify({
        statusg5: statusg5,
        taskidg5: taskidg5, durationg5: durationg5, startg5: startg5, endg5: endg5,
        resourceg5: resourceg5, groupg5: groupg5, targetg5: targetg5, dayg5: dayg5
      }))
      //Group6
      let statusg6 = document.getElementById("datastatusg6").value;
      let taskidg6 = document.getElementById("datataskidg6").value;
      let durationg6 = document.getElementById("datadurationg6").value;
      let startg6 = document.getElementById("datastartg6").value;
      let endg6 = document.getElementById("dataendg6").value;
      let resourceg6 = document.getElementById("dataresourceg6").value;
      let groupg6 = document.getElementById("datagroupg6").value;
      let targetg6 = document.getElementById("targetrowg6").value;
      let dayg6 = document.getElementById("dayg6").value;
      let activeg6 = document.getElementById("activeg6").value;

      localStorage.setItem("g6config", JSON.stringify({
        statusg6: statusg6,
        taskidg6: taskidg6, durationg6: durationg6, startg6: startg6, endg6: endg6,
        resourceg6: resourceg6, groupg6: groupg6, targetg6: targetg6, dayg6: dayg6
      }))
      //Group7
      let statusg7 = document.getElementById("datastatusg7").value;
      let taskidg7 = document.getElementById("datataskidg7").value;
      let durationg7 = document.getElementById("datadurationg7").value;
      let startg7 = document.getElementById("datastartg7").value;
      let endg7 = document.getElementById("dataendg7").value;
      let resourceg7 = document.getElementById("dataresourceg7").value;
      let groupg7 = document.getElementById("datagroupg7").value;
      let targetg7 = document.getElementById("targetrowg7").value;
      let dayg7 = document.getElementById("dayg7").value;
      let activeg7 = document.getElementById("activeg7").value;

      localStorage.setItem("g7config", JSON.stringify({
        statusg7: statusg7,
        taskidg7: taskidg7, durationg7: durationg7, startg7: startg7, endg7: endg7,
        resourceg7: resourceg7, groupg7: groupg7, targetg7: targetg7, dayg7: dayg7
      }))
      //Group8
      let statusg8 = document.getElementById("datastatusg8").value;
      let taskidg8 = document.getElementById("datataskidg8").value;
      let durationg8 = document.getElementById("datadurationg8").value;
      let startg8 = document.getElementById("datastartg8").value;
      let endg8 = document.getElementById("dataendg8").value;
      let resourceg8 = document.getElementById("dataresourceg8").value;
      let groupg8 = document.getElementById("datagroupg8").value;
      let targetg8 = document.getElementById("targetrowg8").value;
      let dayg8 = document.getElementById("dayg8").value;
      let activeg8 = document.getElementById("activeg8").value;

      localStorage.setItem("g8config", JSON.stringify({
        statusg8: statusg8,
        taskidg8: taskidg8, durationg8: durationg8, startg8: startg8, endg8: endg8,
        resourceg8: resourceg8, groupg8: groupg8, targetg8: targetg8, dayg8: dayg8
      }))
      //Group9
      let statusg9 = document.getElementById("datastatusg9").value;
      let taskidg9 = document.getElementById("datataskidg9").value;
      let durationg9 = document.getElementById("datadurationg9").value;
      let startg9 = document.getElementById("datastartg9").value;
      let endg9 = document.getElementById("dataendg9").value;
      let resourceg9 = document.getElementById("dataresourceg9").value;
      let groupg9 = document.getElementById("datagroupg9").value;
      let targetg9 = document.getElementById("targetrowg9").value;
      let dayg9 = document.getElementById("dayg9").value;
      let activeg9 = document.getElementById("activeg9").value;

      localStorage.setItem("g9config", JSON.stringify({
        statusg9: statusg9,
        taskidg9: taskidg9, durationg9: durationg9, startg9: startg9, endg9: endg9,
        resourceg9: resourceg9, groupg9: groupg9, targetg9: targetg9, dayg9: dayg9
      }))
      //Group10
      let statusg10 = document.getElementById("datastatusg10").value;
      let taskidg10 = document.getElementById("datataskidg10").value;
      let durationg10 = document.getElementById("datadurationg10").value;
      let startg10 = document.getElementById("datastartg10").value;
      let endg10 = document.getElementById("dataendg10").value;
      let resourceg10 = document.getElementById("dataresourceg10").value;
      let groupg10 = document.getElementById("datagroupg10").value;
      let targetg10 = document.getElementById("targetrowg10").value;
      let dayg10 = document.getElementById("dayg10").value;
      let activeg10 = document.getElementById("activeg10").value;

      localStorage.setItem("g10config", JSON.stringify({
        statusg10: statusg10,
        taskidg10: taskidg10, durationg10: durationg10, startg10: startg10, endg10: endg10,
        resourceg10: resourceg10, groupg10: groupg10, targetg10: targetg10, dayg10: dayg10
      }))
      //Group11
      let statusg11 = document.getElementById("datastatusg11").value;
      let taskidg11 = document.getElementById("datataskidg11").value;
      let durationg11 = document.getElementById("datadurationg11").value;
      let startg11 = document.getElementById("datastartg11").value;
      let endg11 = document.getElementById("dataendg11").value;
      let resourceg11 = document.getElementById("dataresourceg11").value;
      let groupg11 = document.getElementById("datagroupg11").value;
      let targetg11 = document.getElementById("targetrowg11").value;
      let dayg11 = document.getElementById("dayg11").value;
      let activeg11 = document.getElementById("activeg11").value;

      localStorage.setItem("g11config", JSON.stringify({
        statusg11: statusg11,
        taskidg11: taskidg11, durationg11: durationg11, startg11: startg11, endg11: endg11,
        resourceg11: resourceg11, groupg11: groupg11, targetg11: targetg11, dayg11: dayg11
      }))
      //Group12
      let statusg12 = document.getElementById("datastatusg12").value;
      let taskidg12 = document.getElementById("datataskidg12").value;
      let durationg12 = document.getElementById("datadurationg12").value;
      let startg12 = document.getElementById("datastartg12").value;
      let endg12 = document.getElementById("dataendg12").value;
      let resourceg12 = document.getElementById("dataresourceg12").value;
      let groupg12 = document.getElementById("datagroupg12").value;
      let targetg12 = document.getElementById("targetrowg12").value;
      let dayg12 = document.getElementById("dayg12").value;
      let activeg12 = document.getElementById("activeg12").value;

      localStorage.setItem("g12config", JSON.stringify({
        statusg12: statusg12,
        taskidg12: taskidg12, durationg12: durationg12, startg12: startg12, endg12: endg12,
        resourceg12: resourceg12, groupg12: groupg12, targetg12: targetg12, dayg12: dayg12
      }))
      //Group13
      let statusg13 = document.getElementById("datastatusg13").value;
      let taskidg13 = document.getElementById("datataskidg13").value;
      let durationg13 = document.getElementById("datadurationg13").value;
      let startg13 = document.getElementById("datastartg13").value;
      let endg13 = document.getElementById("dataendg13").value;
      let resourceg13 = document.getElementById("dataresourceg13").value;
      let groupg13 = document.getElementById("datagroupg13").value;
      let targetg13 = document.getElementById("targetrowg13").value;
      let dayg13 = document.getElementById("dayg13").value;
      let activeg13 = document.getElementById("activeg13").value;

      localStorage.setItem("g13config", JSON.stringify({
        statusg13: statusg13,
        taskidg13: taskidg13, durationg13: durationg13, startg13: startg13, endg13: endg13,
        resourceg13: resourceg13, groupg13: groupg13, targetg13: targetg13, dayg13: dayg13
      }))
      //Group14
      let statusg14 = document.getElementById("datastatusg14").value;
      let taskidg14 = document.getElementById("datataskidg14").value;
      let durationg14 = document.getElementById("datadurationg14").value;
      let startg14 = document.getElementById("datastartg14").value;
      let endg14 = document.getElementById("dataendg14").value;
      let resourceg14 = document.getElementById("dataresourceg14").value;
      let groupg14 = document.getElementById("datagroupg14").value;
      let targetg14 = document.getElementById("targetrowg14").value;
      let dayg14 = document.getElementById("dayg14").value;
      let activeg14 = document.getElementById("activeg14").value;

      localStorage.setItem("g14config", JSON.stringify({
        statusg14: statusg14,
        taskidg14: taskidg14, durationg14: durationg14, startg14: startg14, endg14: endg14,
        resourceg14: resourceg14, groupg14: groupg14, targetg14: targetg14, dayg14: dayg14
      }))
      //Group15
      let statusg15 = document.getElementById("datastatusg15").value;
      let taskidg15 = document.getElementById("datataskidg15").value;
      let durationg15 = document.getElementById("datadurationg15").value;
      let startg15 = document.getElementById("datastartg15").value;
      let endg15 = document.getElementById("dataendg15").value;
      let resourceg15 = document.getElementById("dataresourceg15").value;
      let groupg15 = document.getElementById("datagroupg15").value;
      let targetg15 = document.getElementById("targetrowg15").value;
      let dayg15 = document.getElementById("dayg15").value;
      let activeg15 = document.getElementById("activeg15").value;

      localStorage.setItem("g15config", JSON.stringify({
        statusg15: statusg15,
        taskidg15: taskidg15, durationg15: durationg15, startg15: startg15, endg15: endg15,
        resourceg15: resourceg15, groupg15: groupg15, targetg15: targetg15, dayg15: dayg15
      }))
      //Group16
      let statusg16 = document.getElementById("datastatusg16").value;
      let taskidg16 = document.getElementById("datataskidg16").value;
      let durationg16 = document.getElementById("datadurationg16").value;
      let startg16 = document.getElementById("datastartg16").value;
      let endg16 = document.getElementById("dataendg16").value;
      let resourceg16 = document.getElementById("dataresourceg16").value;
      let groupg16 = document.getElementById("datagroupg16").value;
      let targetg16 = document.getElementById("targetrowg16").value;
      let dayg16 = document.getElementById("dayg16").value;
      let activeg16 = document.getElementById("activeg16").value;

      localStorage.setItem("g16config", JSON.stringify({
        statusg16: statusg16,
        taskidg16: taskidg16, durationg16: durationg16, startg16: startg16, endg16: endg16,
        resourceg16: resourceg16, groupg16: groupg16, targetg16: targetg16, dayg16: dayg16
      }))
      //Group17
      let statusg17 = document.getElementById("datastatusg17").value;
      let taskidg17 = document.getElementById("datataskidg17").value;
      let durationg17 = document.getElementById("datadurationg17").value;
      let startg17 = document.getElementById("datastartg17").value;
      let endg17 = document.getElementById("dataendg17").value;
      let resourceg17 = document.getElementById("dataresourceg17").value;
      let groupg17 = document.getElementById("datagroupg17").value;
      let targetg17 = document.getElementById("targetrowg17").value;
      let dayg17 = document.getElementById("dayg17").value;
      let activeg17 = document.getElementById("activeg17").value;

      localStorage.setItem("g17config", JSON.stringify({
        statusg17: statusg17,
        taskidg17: taskidg17, durationg17: durationg17, startg17: startg17, endg17: endg17,
        resourceg17: resourceg17, groupg17: groupg17, targetg17: targetg17, dayg17: dayg17
      }))
      //Group18
      let statusg18 = document.getElementById("datastatusg18").value;
      let taskidg18 = document.getElementById("datataskidg18").value;
      let durationg18 = document.getElementById("datadurationg18").value;
      let startg18 = document.getElementById("datastartg18").value;
      let endg18 = document.getElementById("dataendg18").value;
      let resourceg18 = document.getElementById("dataresourceg18").value;
      let groupg18 = document.getElementById("datagroupg18").value;
      let targetg18 = document.getElementById("targetrowg18").value;
      let dayg18 = document.getElementById("dayg18").value;
      let activeg18 = document.getElementById("activeg18").value;

      localStorage.setItem("g18config", JSON.stringify({
        statusg18: statusg18,
        taskidg18: taskidg18, durationg18: durationg18, startg18: startg18, endg18: endg18,
        resourceg18: resourceg18, groupg18: groupg18, targetg18: targetg18, dayg18: dayg18
      }))
      //Group19
      let statusg19 = document.getElementById("datastatusg19").value;
      let taskidg19 = document.getElementById("datataskidg19").value;
      let durationg19 = document.getElementById("datadurationg19").value;
      let startg19 = document.getElementById("datastartg19").value;
      let endg19 = document.getElementById("dataendg19").value;
      let resourceg19 = document.getElementById("dataresourceg19").value;
      let groupg19 = document.getElementById("datagroupg19").value;
      let targetg19 = document.getElementById("targetrowg19").value;
      let dayg19 = document.getElementById("dayg19").value;
      let activeg19 = document.getElementById("activeg19").value;

      localStorage.setItem("g19config", JSON.stringify({
        statusg19: statusg19,
        taskidg19: taskidg19, durationg19: durationg19, startg19: startg19, endg19: endg19,
        resourceg19: resourceg19, groupg19: groupg19, targetg19: targetg19, dayg19: dayg19
      }))
      //Group20
      let statusg20 = document.getElementById("datastatusg20").value;
      let taskidg20 = document.getElementById("datataskidg20").value;
      let durationg20 = document.getElementById("datadurationg20").value;
      let startg20 = document.getElementById("datastartg20").value;
      let endg20 = document.getElementById("dataendg20").value;
      let resourceg20 = document.getElementById("dataresourceg20").value;
      let groupg20 = document.getElementById("datagroupg20").value;
      let targetg20 = document.getElementById("targetrowg20").value;
      let dayg20 = document.getElementById("dayg20").value;
      let activeg20 = document.getElementById("activeg20").value;

      localStorage.setItem("g20config", JSON.stringify({
        statusg20: statusg20,
        taskidg20: taskidg20, durationg20: durationg20, startg20: startg20, endg20: endg20,
        resourceg20: resourceg20, groupg20: groupg20, targetg20: targetg20, dayg20: dayg20
      }))
      //Group21
      let statusg21 = document.getElementById("datastatusg21").value;
      let taskidg21 = document.getElementById("datataskidg21").value;
      let durationg21 = document.getElementById("datadurationg21").value;
      let startg21 = document.getElementById("datastartg21").value;
      let endg21 = document.getElementById("dataendg21").value;
      let resourceg21 = document.getElementById("dataresourceg21").value;
      let groupg21 = document.getElementById("datagroupg21").value;
      let targetg21 = document.getElementById("targetrowg21").value;
      let dayg21 = document.getElementById("dayg21").value;
      let activeg21 = document.getElementById("activeg21").value;

      localStorage.setItem("g21config", JSON.stringify({
        statusg21: statusg21,
        taskidg21: taskidg21, durationg21: durationg21, startg21: startg21, endg21: endg21,
        resourceg21: resourceg21, groupg21: groupg21, targetg21: targetg21, dayg21: dayg21
      }))
      //Group22
      let statusg22 = document.getElementById("datastatusg22").value;
      let taskidg22 = document.getElementById("datataskidg22").value;
      let durationg22 = document.getElementById("datadurationg22").value;
      let startg22 = document.getElementById("datastartg22").value;
      let endg22 = document.getElementById("dataendg22").value;
      let resourceg22 = document.getElementById("dataresourceg22").value;
      let groupg22 = document.getElementById("datagroupg22").value;
      let targetg22 = document.getElementById("targetrowg22").value;
      let dayg22 = document.getElementById("dayg22").value;
      let activeg22 = document.getElementById("activeg22").value;

      localStorage.setItem("g22config", JSON.stringify({
        statusg22: statusg22,
        taskidg22: taskidg22, durationg22: durationg22, startg22: startg22, endg22: endg22,
        resourceg22: resourceg22, groupg22: groupg22, targetg22: targetg22, dayg22: dayg22
      }))
      //Group23
      let statusg23 = document.getElementById("datastatusg23").value;
      let taskidg23 = document.getElementById("datataskidg23").value;
      let durationg23 = document.getElementById("datadurationg23").value;
      let startg23 = document.getElementById("datastartg23").value;
      let endg23 = document.getElementById("dataendg23").value;
      let resourceg23 = document.getElementById("dataresourceg23").value;
      let groupg23 = document.getElementById("datagroupg23").value;
      let targetg23 = document.getElementById("targetrowg23").value;
      let dayg23 = document.getElementById("dayg23").value;
      let activeg23 = document.getElementById("activeg23").value;

      localStorage.setItem("g23config", JSON.stringify({
        statusg23: statusg23,
        taskidg23: taskidg23, durationg23: durationg23, startg23: startg23, endg23: endg23,
        resourceg23: resourceg23, groupg23: groupg23, targetg23: targetg23, dayg23: dayg23
      }))
      //Group24
      let statusg24 = document.getElementById("datastatusg24").value;
      let taskidg24 = document.getElementById("datataskidg24").value;
      let durationg24 = document.getElementById("datadurationg24").value;
      let startg24 = document.getElementById("datastartg24").value;
      let endg24 = document.getElementById("dataendg24").value;
      let resourceg24 = document.getElementById("dataresourceg24").value;
      let groupg24 = document.getElementById("datagroupg24").value;
      let targetg24 = document.getElementById("targetrowg24").value;
      let dayg24 = document.getElementById("dayg24").value;
      let activeg24 = document.getElementById("activeg24").value;

      localStorage.setItem("g24config", JSON.stringify({
        statusg24: statusg24,
        taskidg24: taskidg24, durationg24: durationg24, startg24: startg24, endg24: endg24,
        resourceg24: resourceg24, groupg24: groupg24, targetg24: targetg24, dayg24: dayg24
      }))
      //Group25
      let statusg25 = document.getElementById("datastatusg25").value;
      let taskidg25 = document.getElementById("datataskidg25").value;
      let durationg25 = document.getElementById("datadurationg25").value;
      let startg25 = document.getElementById("datastartg25").value;
      let endg25 = document.getElementById("dataendg25").value;
      let resourceg25 = document.getElementById("dataresourceg25").value;
      let groupg25 = document.getElementById("datagroupg25").value;
      let targetg25 = document.getElementById("targetrowg25").value;
      let dayg25 = document.getElementById("dayg25").value;
      let activeg25 = document.getElementById("activeg25").value;

      localStorage.setItem("g25config", JSON.stringify({
        statusg25: statusg25,
        taskidg25: taskidg25, durationg25: durationg25, startg25: startg25, endg25: endg25,
        resourceg25: resourceg25, groupg25: groupg25, targetg25: targetg25, dayg25: dayg25
      }))
      //Group26
      let statusg26 = document.getElementById("datastatusg26").value;
      let taskidg26 = document.getElementById("datataskidg26").value;
      let durationg26 = document.getElementById("datadurationg26").value;
      let startg26 = document.getElementById("datastartg26").value;
      let endg26 = document.getElementById("dataendg26").value;
      let resourceg26 = document.getElementById("dataresourceg26").value;
      let groupg26 = document.getElementById("datagroupg26").value;
      let targetg26 = document.getElementById("targetrowg26").value;
      let dayg26 = document.getElementById("dayg26").value;
      let activeg26 = document.getElementById("activeg26").value;

      localStorage.setItem("g26config", JSON.stringify({
        statusg26: statusg26,
        taskidg26: taskidg26, durationg26: durationg26, startg26: startg26, endg26: endg26,
        resourceg26: resourceg26, groupg26: groupg26, targetg26: targetg26, dayg26: dayg26
      }))
      //Group27
      let statusg27 = document.getElementById("datastatusg27").value;
      let taskidg27 = document.getElementById("datataskidg27").value;
      let durationg27 = document.getElementById("datadurationg27").value;
      let startg27 = document.getElementById("datastartg27").value;
      let endg27 = document.getElementById("dataendg27").value;
      let resourceg27 = document.getElementById("dataresourceg27").value;
      let groupg27 = document.getElementById("datagroupg27").value;
      let targetg27 = document.getElementById("targetrowg27").value;
      let dayg27 = document.getElementById("dayg27").value;
      let activeg27 = document.getElementById("activeg27").value;

      localStorage.setItem("g27config", JSON.stringify({
        statusg27: statusg27,
        taskidg27: taskidg27, durationg27: durationg27, startg27: startg27, endg27: endg27,
        resourceg27: resourceg27, groupg27: groupg27, targetg27: targetg27, dayg27: dayg27
      }))
      //Group28
      let statusg28 = document.getElementById("datastatusg28").value;
      let taskidg28 = document.getElementById("datataskidg28").value;
      let durationg28 = document.getElementById("datadurationg28").value;
      let startg28 = document.getElementById("datastartg28").value;
      let endg28 = document.getElementById("dataendg28").value;
      let resourceg28 = document.getElementById("dataresourceg28").value;
      let groupg28 = document.getElementById("datagroupg28").value;
      let targetg28 = document.getElementById("targetrowg28").value;
      let dayg28 = document.getElementById("dayg28").value;
      let activeg28 = document.getElementById("activeg28").value;

      localStorage.setItem("g28config", JSON.stringify({
        statusg28: statusg28,
        taskidg28: taskidg28, durationg28: durationg28, startg28: startg28, endg28: endg28,
        resourceg28: resourceg28, groupg28: groupg28, targetg28: targetg28, dayg28: dayg28
      }))
      //Group29
      let statusg29 = document.getElementById("datastatusg29").value;
      let taskidg29 = document.getElementById("datataskidg29").value;
      let durationg29 = document.getElementById("datadurationg29").value;
      let startg29 = document.getElementById("datastartg29").value;
      let endg29 = document.getElementById("dataendg29").value;
      let resourceg29 = document.getElementById("dataresourceg29").value;
      let groupg29 = document.getElementById("datagroupg29").value;
      let targetg29 = document.getElementById("targetrowg29").value;
      let dayg29 = document.getElementById("dayg29").value;
      let activeg29 = document.getElementById("activeg29").value;

      localStorage.setItem("g29config", JSON.stringify({
        statusg29: statusg29,
        taskidg29: taskidg29, durationg29: durationg29, startg29: startg29, endg29: endg29,
        resourceg29: resourceg29, groupg29: groupg29, targetg29: targetg29, dayg29: dayg29
      }))
      //Group30
      let statusg30 = document.getElementById("datastatusg30").value;
      let taskidg30 = document.getElementById("datataskidg30").value;
      let durationg30 = document.getElementById("datadurationg30").value;
      let startg30 = document.getElementById("datastartg30").value;
      let endg30 = document.getElementById("dataendg30").value;
      let resourceg30 = document.getElementById("dataresourceg30").value;
      let groupg30 = document.getElementById("datagroupg30").value;
      let targetg30 = document.getElementById("targetrowg30").value;
      let dayg30 = document.getElementById("dayg30").value;
      let activeg30 = document.getElementById("activeg30").value;

      localStorage.setItem("g30config", JSON.stringify({
        statusg30: statusg30,
        taskidg30: taskidg30, durationg30: durationg30, startg30: startg30, endg30: endg30,
        resourceg30: resourceg30, groupg30: groupg30, targetg30: targetg30, dayg30: dayg30
      }))

      try {
        await Excel.run(async (context) => {
          // console.log(sname)
          let sheet1 = context.workbook.worksheets.getItem(sname);
          let sheet2 = context.workbook.worksheets.getItem(tname);
          let inprogresstask = sheet2.getRange(inprogressaddress)
          let donetask = sheet2.getRange(completeaddress)
          let delaytask = sheet2.getRange(delayaddress)
          let ICtask = sheet2.getRange(Icompleteaddress)
          let CDtask = sheet2.getRange(Cdelayaddress)

          //Group1
          let statusrangeg1 = sheet1.getRange(statusg1);
          let taskrangeg1 = sheet1.getRange(taskidg1);
          let durationrangeg1 = sheet1.getRange(durationg1);
          let startrangeg1 = sheet1.getRange(startg1);
          let endrangeg1 = sheet1.getRange(endg1);
          let resourcerangeg1 = sheet1.getRange(resourceg1);
          let grouprangeg1 = sheet1.getRange(groupg1);
          statusrangeg1.load(["text", "address"])
          taskrangeg1.load(["text", "address"])
          durationrangeg1.load(["text", "address"])
          startrangeg1.load(["text", "address"])
          endrangeg1.load(["text", "address"])
          resourcerangeg1.load(["text", "address"])
          grouprangeg1.load(["text", "address"])
          //Group2
          let statusrangeg2 = sheet1.getRange(statusg2);
          let taskrangeg2 = sheet1.getRange(taskidg2);
          let durationrangeg2 = sheet1.getRange(durationg2);
          let startrangeg2 = sheet1.getRange(startg2);
          let endrangeg2 = sheet1.getRange(endg2);
          let resourcerangeg2 = sheet1.getRange(resourceg2);
          let grouprangeg2 = sheet1.getRange(groupg2);
          statusrangeg2.load(["text", "address"])
          taskrangeg2.load(["text", "address"])
          durationrangeg2.load(["text", "address"])
          startrangeg2.load(["text", "address"])
          endrangeg2.load(["text", "address"])
          resourcerangeg2.load(["text", "address"])
          grouprangeg2.load(["text", "address"])
          //Group3
          let statusrangeg3 = sheet1.getRange(statusg3);
          let taskrangeg3 = sheet1.getRange(taskidg3);
          let durationrangeg3 = sheet1.getRange(durationg3);
          let startrangeg3 = sheet1.getRange(startg3);
          let endrangeg3 = sheet1.getRange(endg3);
          let resourcerangeg3 = sheet1.getRange(resourceg3);
          let grouprangeg3 = sheet1.getRange(groupg3);
          statusrangeg3.load(["text", "address"])
          taskrangeg3.load(["text", "address"])
          durationrangeg3.load(["text", "address"])
          startrangeg3.load(["text", "address"])
          endrangeg3.load(["text", "address"])
          resourcerangeg3.load(["text", "address"])
          grouprangeg3.load(["text", "address"])
          //Group4
          let statusrangeg4 = sheet1.getRange(statusg4);
          let taskrangeg4 = sheet1.getRange(taskidg4);
          let durationrangeg4 = sheet1.getRange(durationg4);
          let startrangeg4 = sheet1.getRange(startg4);
          let endrangeg4 = sheet1.getRange(endg4);
          let resourcerangeg4 = sheet1.getRange(resourceg4);
          let grouprangeg4 = sheet1.getRange(groupg4);
          statusrangeg4.load(["text", "address"])
          taskrangeg4.load(["text", "address"])
          durationrangeg4.load(["text", "address"])
          startrangeg4.load(["text", "address"])
          endrangeg4.load(["text", "address"])
          resourcerangeg4.load(["text", "address"])
          grouprangeg4.load(["text", "address"])
          //Group5
          let statusrangeg5 = sheet1.getRange(statusg5);
          let taskrangeg5 = sheet1.getRange(taskidg5);
          let durationrangeg5 = sheet1.getRange(durationg5);
          let startrangeg5 = sheet1.getRange(startg5);
          let endrangeg5 = sheet1.getRange(endg5);
          let resourcerangeg5 = sheet1.getRange(resourceg5);
          let grouprangeg5 = sheet1.getRange(groupg5);
          statusrangeg5.load(["text", "address"])
          taskrangeg5.load(["text", "address"])
          durationrangeg5.load(["text", "address"])
          startrangeg5.load(["text", "address"])
          endrangeg5.load(["text", "address"])
          resourcerangeg5.load(["text", "address"])
          grouprangeg5.load(["text", "address"])
          //Group6
          let statusrangeg6 = sheet1.getRange(statusg6);
          let taskrangeg6 = sheet1.getRange(taskidg6);
          let durationrangeg6 = sheet1.getRange(durationg6);
          let startrangeg6 = sheet1.getRange(startg6);
          let endrangeg6 = sheet1.getRange(endg6);
          let resourcerangeg6 = sheet1.getRange(resourceg6);
          let grouprangeg6 = sheet1.getRange(groupg6);
          statusrangeg6.load(["text", "address"])
          taskrangeg6.load(["text", "address"])
          durationrangeg6.load(["text", "address"])
          startrangeg6.load(["text", "address"])
          endrangeg6.load(["text", "address"])
          resourcerangeg6.load(["text", "address"])
          grouprangeg6.load(["text", "address"])
          //Group7
          let statusrangeg7 = sheet1.getRange(statusg7);
          let taskrangeg7 = sheet1.getRange(taskidg7);
          let durationrangeg7 = sheet1.getRange(durationg7);
          let startrangeg7 = sheet1.getRange(startg7);
          let endrangeg7 = sheet1.getRange(endg7);
          let resourcerangeg7 = sheet1.getRange(resourceg7);
          let grouprangeg7 = sheet1.getRange(groupg7);
          statusrangeg7.load(["text", "address"])
          taskrangeg7.load(["text", "address"])
          durationrangeg7.load(["text", "address"])
          startrangeg7.load(["text", "address"])
          endrangeg7.load(["text", "address"])
          resourcerangeg7.load(["text", "address"])
          grouprangeg7.load(["text", "address"])
          //Group8
          let statusrangeg8 = sheet1.getRange(statusg8);
          let taskrangeg8 = sheet1.getRange(taskidg8);
          let durationrangeg8 = sheet1.getRange(durationg8);
          let startrangeg8 = sheet1.getRange(startg8);
          let endrangeg8 = sheet1.getRange(endg8);
          let resourcerangeg8 = sheet1.getRange(resourceg8);
          let grouprangeg8 = sheet1.getRange(groupg8);
          statusrangeg8.load(["text", "address"])
          taskrangeg8.load(["text", "address"])
          durationrangeg8.load(["text", "address"])
          startrangeg8.load(["text", "address"])
          endrangeg8.load(["text", "address"])
          resourcerangeg8.load(["text", "address"])
          grouprangeg8.load(["text", "address"])
          //Group9
          let statusrangeg9 = sheet1.getRange(statusg9);
          let taskrangeg9 = sheet1.getRange(taskidg9);
          let durationrangeg9 = sheet1.getRange(durationg9);
          let startrangeg9 = sheet1.getRange(startg9);
          let endrangeg9 = sheet1.getRange(endg9);
          let resourcerangeg9 = sheet1.getRange(resourceg9);
          let grouprangeg9 = sheet1.getRange(groupg9);
          statusrangeg9.load(["text", "address"])
          taskrangeg9.load(["text", "address"])
          durationrangeg9.load(["text", "address"])
          startrangeg9.load(["text", "address"])
          endrangeg9.load(["text", "address"])
          resourcerangeg9.load(["text", "address"])
          grouprangeg9.load(["text", "address"])
          //Group10
          let statusrangeg10 = sheet1.getRange(statusg10);
          let taskrangeg10 = sheet1.getRange(taskidg10);
          let durationrangeg10 = sheet1.getRange(durationg10);
          let startrangeg10 = sheet1.getRange(startg10);
          let endrangeg10 = sheet1.getRange(endg10);
          let resourcerangeg10 = sheet1.getRange(resourceg10);
          let grouprangeg10 = sheet1.getRange(groupg10);
          statusrangeg10.load(["text", "address"])
          taskrangeg10.load(["text", "address"])
          durationrangeg10.load(["text", "address"])
          startrangeg10.load(["text", "address"])
          endrangeg10.load(["text", "address"])
          resourcerangeg10.load(["text", "address"])
          grouprangeg10.load(["text", "address"])
          //Group11
          let statusrangeg11 = sheet1.getRange(statusg11);
          let taskrangeg11 = sheet1.getRange(taskidg11);
          let durationrangeg11 = sheet1.getRange(durationg11);
          let startrangeg11 = sheet1.getRange(startg11);
          let endrangeg11 = sheet1.getRange(endg11);
          let resourcerangeg11 = sheet1.getRange(resourceg11);
          let grouprangeg11 = sheet1.getRange(groupg11);
          statusrangeg11.load(["text", "address"])
          taskrangeg11.load(["text", "address"])
          durationrangeg11.load(["text", "address"])
          startrangeg11.load(["text", "address"])
          endrangeg11.load(["text", "address"])
          resourcerangeg11.load(["text", "address"])
          grouprangeg11.load(["text", "address"])
          //Group12
          let statusrangeg12 = sheet1.getRange(statusg12);
          let taskrangeg12 = sheet1.getRange(taskidg12);
          let durationrangeg12 = sheet1.getRange(durationg12);
          let startrangeg12 = sheet1.getRange(startg12);
          let endrangeg12 = sheet1.getRange(endg12);
          let resourcerangeg12 = sheet1.getRange(resourceg12);
          let grouprangeg12 = sheet1.getRange(groupg12);
          statusrangeg12.load(["text", "address"])
          taskrangeg12.load(["text", "address"])
          durationrangeg12.load(["text", "address"])
          startrangeg12.load(["text", "address"])
          endrangeg12.load(["text", "address"])
          resourcerangeg12.load(["text", "address"])
          grouprangeg12.load(["text", "address"])
          //Group13
          let statusrangeg13 = sheet1.getRange(statusg13);
          let taskrangeg13 = sheet1.getRange(taskidg13);
          let durationrangeg13 = sheet1.getRange(durationg13);
          let startrangeg13 = sheet1.getRange(startg13);
          let endrangeg13 = sheet1.getRange(endg13);
          let resourcerangeg13 = sheet1.getRange(resourceg13);
          let grouprangeg13 = sheet1.getRange(groupg13);
          statusrangeg13.load(["text", "address"])
          taskrangeg13.load(["text", "address"])
          durationrangeg13.load(["text", "address"])
          startrangeg13.load(["text", "address"])
          endrangeg13.load(["text", "address"])
          resourcerangeg13.load(["text", "address"])
          grouprangeg13.load(["text", "address"])
          //Group14
          let statusrangeg14 = sheet1.getRange(statusg14);
          let taskrangeg14 = sheet1.getRange(taskidg14);
          let durationrangeg14 = sheet1.getRange(durationg14);
          let startrangeg14 = sheet1.getRange(startg14);
          let endrangeg14 = sheet1.getRange(endg14);
          let resourcerangeg14 = sheet1.getRange(resourceg14);
          let grouprangeg14 = sheet1.getRange(groupg14);
          statusrangeg14.load(["text", "address"])
          taskrangeg14.load(["text", "address"])
          durationrangeg14.load(["text", "address"])
          startrangeg14.load(["text", "address"])
          endrangeg14.load(["text", "address"])
          resourcerangeg14.load(["text", "address"])
          grouprangeg14.load(["text", "address"])
          //Group15
          let statusrangeg15 = sheet1.getRange(statusg15);
          let taskrangeg15 = sheet1.getRange(taskidg15);
          let durationrangeg15 = sheet1.getRange(durationg15);
          let startrangeg15 = sheet1.getRange(startg15);
          let endrangeg15 = sheet1.getRange(endg15);
          let resourcerangeg15 = sheet1.getRange(resourceg15);
          let grouprangeg15 = sheet1.getRange(groupg15);
          statusrangeg15.load(["text", "address"])
          taskrangeg15.load(["text", "address"])
          durationrangeg15.load(["text", "address"])
          startrangeg15.load(["text", "address"])
          endrangeg15.load(["text", "address"])
          resourcerangeg15.load(["text", "address"])
          grouprangeg15.load(["text", "address"])
          //Group16
          let statusrangeg16 = sheet1.getRange(statusg16);
          let taskrangeg16 = sheet1.getRange(taskidg16);
          let durationrangeg16 = sheet1.getRange(durationg16);
          let startrangeg16 = sheet1.getRange(startg16);
          let endrangeg16 = sheet1.getRange(endg16);
          let resourcerangeg16 = sheet1.getRange(resourceg16);
          let grouprangeg16 = sheet1.getRange(groupg16);
          statusrangeg16.load(["text", "address"])
          taskrangeg16.load(["text", "address"])
          durationrangeg16.load(["text", "address"])
          startrangeg16.load(["text", "address"])
          endrangeg16.load(["text", "address"])
          resourcerangeg16.load(["text", "address"])
          grouprangeg16.load(["text", "address"])
          //Group17
          let statusrangeg17 = sheet1.getRange(statusg17);
          let taskrangeg17 = sheet1.getRange(taskidg17);
          let durationrangeg17 = sheet1.getRange(durationg17);
          let startrangeg17 = sheet1.getRange(startg17);
          let endrangeg17 = sheet1.getRange(endg17);
          let resourcerangeg17 = sheet1.getRange(resourceg17);
          let grouprangeg17 = sheet1.getRange(groupg17);
          statusrangeg17.load(["text", "address"])
          taskrangeg17.load(["text", "address"])
          durationrangeg17.load(["text", "address"])
          startrangeg17.load(["text", "address"])
          endrangeg17.load(["text", "address"])
          resourcerangeg17.load(["text", "address"])
          grouprangeg17.load(["text", "address"])
          //Group18
          let statusrangeg18 = sheet1.getRange(statusg18);
          let taskrangeg18 = sheet1.getRange(taskidg18);
          let durationrangeg18 = sheet1.getRange(durationg18);
          let startrangeg18 = sheet1.getRange(startg18);
          let endrangeg18 = sheet1.getRange(endg18);
          let resourcerangeg18 = sheet1.getRange(resourceg18);
          let grouprangeg18 = sheet1.getRange(groupg18);
          statusrangeg18.load(["text", "address"])
          taskrangeg18.load(["text", "address"])
          durationrangeg18.load(["text", "address"])
          startrangeg18.load(["text", "address"])
          endrangeg18.load(["text", "address"])
          resourcerangeg18.load(["text", "address"])
          grouprangeg18.load(["text", "address"])
          //Group19
          let statusrangeg19 = sheet1.getRange(statusg19);
          let taskrangeg19 = sheet1.getRange(taskidg19);
          let durationrangeg19 = sheet1.getRange(durationg19);
          let startrangeg19 = sheet1.getRange(startg19);
          let endrangeg19 = sheet1.getRange(endg19);
          let resourcerangeg19 = sheet1.getRange(resourceg19);
          let grouprangeg19 = sheet1.getRange(groupg19);
          statusrangeg19.load(["text", "address"])
          taskrangeg19.load(["text", "address"])
          durationrangeg19.load(["text", "address"])
          startrangeg19.load(["text", "address"])
          endrangeg19.load(["text", "address"])
          resourcerangeg19.load(["text", "address"])
          grouprangeg19.load(["text", "address"])
          //Group20
          let statusrangeg20 = sheet1.getRange(statusg20);
          let taskrangeg20 = sheet1.getRange(taskidg20);
          let durationrangeg20 = sheet1.getRange(durationg20);
          let startrangeg20 = sheet1.getRange(startg20);
          let endrangeg20 = sheet1.getRange(endg20);
          let resourcerangeg20 = sheet1.getRange(resourceg20);
          let grouprangeg20 = sheet1.getRange(groupg20);
          statusrangeg20.load(["text", "address"])
          taskrangeg20.load(["text", "address"])
          durationrangeg20.load(["text", "address"])
          startrangeg20.load(["text", "address"])
          endrangeg20.load(["text", "address"])
          resourcerangeg20.load(["text", "address"])
          grouprangeg20.load(["text", "address"])
          //Group21
          let statusrangeg21 = sheet1.getRange(statusg21);
          let taskrangeg21 = sheet1.getRange(taskidg21);
          let durationrangeg21 = sheet1.getRange(durationg21);
          let startrangeg21 = sheet1.getRange(startg21);
          let endrangeg21 = sheet1.getRange(endg21);
          let resourcerangeg21 = sheet1.getRange(resourceg21);
          let grouprangeg21 = sheet1.getRange(groupg21);
          statusrangeg21.load(["text", "address"])
          taskrangeg21.load(["text", "address"])
          durationrangeg21.load(["text", "address"])
          startrangeg21.load(["text", "address"])
          endrangeg21.load(["text", "address"])
          resourcerangeg21.load(["text", "address"])
          grouprangeg21.load(["text", "address"])
          //Group22
          let statusrangeg22 = sheet1.getRange(statusg22);
          let taskrangeg22 = sheet1.getRange(taskidg22);
          let durationrangeg22 = sheet1.getRange(durationg22);
          let startrangeg22 = sheet1.getRange(startg22);
          let endrangeg22 = sheet1.getRange(endg22);
          let resourcerangeg22 = sheet1.getRange(resourceg22);
          let grouprangeg22 = sheet1.getRange(groupg22);
          statusrangeg22.load(["text", "address"])
          taskrangeg22.load(["text", "address"])
          durationrangeg22.load(["text", "address"])
          startrangeg22.load(["text", "address"])
          endrangeg22.load(["text", "address"])
          resourcerangeg22.load(["text", "address"])
          grouprangeg22.load(["text", "address"])
          //Group23
          let statusrangeg23 = sheet1.getRange(statusg23);
          let taskrangeg23 = sheet1.getRange(taskidg23);
          let durationrangeg23 = sheet1.getRange(durationg23);
          let startrangeg23 = sheet1.getRange(startg23);
          let endrangeg23 = sheet1.getRange(endg23);
          let resourcerangeg23 = sheet1.getRange(resourceg23);
          let grouprangeg23 = sheet1.getRange(groupg23);
          statusrangeg23.load(["text", "address"])
          taskrangeg23.load(["text", "address"])
          durationrangeg23.load(["text", "address"])
          startrangeg23.load(["text", "address"])
          endrangeg23.load(["text", "address"])
          resourcerangeg23.load(["text", "address"])
          grouprangeg23.load(["text", "address"])
          //Group24
          let statusrangeg24 = sheet1.getRange(statusg24);
          let taskrangeg24 = sheet1.getRange(taskidg24);
          let durationrangeg24 = sheet1.getRange(durationg24);
          let startrangeg24 = sheet1.getRange(startg24);
          let endrangeg24 = sheet1.getRange(endg24);
          let resourcerangeg24 = sheet1.getRange(resourceg24);
          let grouprangeg24 = sheet1.getRange(groupg24);
          statusrangeg24.load(["text", "address"])
          taskrangeg24.load(["text", "address"])
          durationrangeg24.load(["text", "address"])
          startrangeg24.load(["text", "address"])
          endrangeg24.load(["text", "address"])
          resourcerangeg24.load(["text", "address"])
          grouprangeg24.load(["text", "address"])
          //Group25
          let statusrangeg25 = sheet1.getRange(statusg25);
          let taskrangeg25 = sheet1.getRange(taskidg25);
          let durationrangeg25 = sheet1.getRange(durationg25);
          let startrangeg25 = sheet1.getRange(startg25);
          let endrangeg25 = sheet1.getRange(endg25);
          let resourcerangeg25 = sheet1.getRange(resourceg25);
          let grouprangeg25 = sheet1.getRange(groupg25);
          statusrangeg25.load(["text", "address"])
          taskrangeg25.load(["text", "address"])
          durationrangeg25.load(["text", "address"])
          startrangeg25.load(["text", "address"])
          endrangeg25.load(["text", "address"])
          resourcerangeg25.load(["text", "address"])
          grouprangeg25.load(["text", "address"])
          //Group26
          let statusrangeg26 = sheet1.getRange(statusg26);
          let taskrangeg26 = sheet1.getRange(taskidg26);
          let durationrangeg26 = sheet1.getRange(durationg26);
          let startrangeg26 = sheet1.getRange(startg26);
          let endrangeg26 = sheet1.getRange(endg26);
          let resourcerangeg26 = sheet1.getRange(resourceg26);
          let grouprangeg26 = sheet1.getRange(groupg26);
          statusrangeg26.load(["text", "address"])
          taskrangeg26.load(["text", "address"])
          durationrangeg26.load(["text", "address"])
          startrangeg26.load(["text", "address"])
          endrangeg26.load(["text", "address"])
          resourcerangeg26.load(["text", "address"])
          grouprangeg26.load(["text", "address"])
          //Group27
          let statusrangeg27 = sheet1.getRange(statusg27);
          let taskrangeg27 = sheet1.getRange(taskidg27);
          let durationrangeg27 = sheet1.getRange(durationg27);
          let startrangeg27 = sheet1.getRange(startg27);
          let endrangeg27 = sheet1.getRange(endg27);
          let resourcerangeg27 = sheet1.getRange(resourceg27);
          let grouprangeg27 = sheet1.getRange(groupg27);
          statusrangeg27.load(["text", "address"])
          taskrangeg27.load(["text", "address"])
          durationrangeg27.load(["text", "address"])
          startrangeg27.load(["text", "address"])
          endrangeg27.load(["text", "address"])
          resourcerangeg27.load(["text", "address"])
          grouprangeg27.load(["text", "address"])
          //Group28
          let statusrangeg28 = sheet1.getRange(statusg28);
          let taskrangeg28 = sheet1.getRange(taskidg28);
          let durationrangeg28 = sheet1.getRange(durationg28);
          let startrangeg28 = sheet1.getRange(startg28);
          let endrangeg28 = sheet1.getRange(endg28);
          let resourcerangeg28 = sheet1.getRange(resourceg28);
          let grouprangeg28 = sheet1.getRange(groupg28);
          statusrangeg28.load(["text", "address"])
          taskrangeg28.load(["text", "address"])
          durationrangeg28.load(["text", "address"])
          startrangeg28.load(["text", "address"])
          endrangeg28.load(["text", "address"])
          resourcerangeg28.load(["text", "address"])
          grouprangeg28.load(["text", "address"])
          //Group29
          let statusrangeg29 = sheet1.getRange(statusg29);
          let taskrangeg29 = sheet1.getRange(taskidg29);
          let durationrangeg29 = sheet1.getRange(durationg29);
          let startrangeg29 = sheet1.getRange(startg29);
          let endrangeg29 = sheet1.getRange(endg29);
          let resourcerangeg29 = sheet1.getRange(resourceg29);
          let grouprangeg29 = sheet1.getRange(groupg29);
          statusrangeg29.load(["text", "address"])
          taskrangeg29.load(["text", "address"])
          durationrangeg29.load(["text", "address"])
          startrangeg29.load(["text", "address"])
          endrangeg29.load(["text", "address"])
          resourcerangeg29.load(["text", "address"])
          grouprangeg29.load(["text", "address"])
          //Group30
          let statusrangeg30 = sheet1.getRange(statusg30);
          let taskrangeg30 = sheet1.getRange(taskidg30);
          let durationrangeg30 = sheet1.getRange(durationg30);
          let startrangeg30 = sheet1.getRange(startg30);
          let endrangeg30 = sheet1.getRange(endg30);
          let resourcerangeg30 = sheet1.getRange(resourceg30);
          let grouprangeg30 = sheet1.getRange(groupg30);
          statusrangeg30.load(["text", "address"])
          taskrangeg30.load(["text", "address"])
          durationrangeg30.load(["text", "address"])
          startrangeg30.load(["text", "address"])
          endrangeg30.load(["text", "address"])
          resourcerangeg30.load(["text", "address"])
          grouprangeg30.load(["text", "address"])

          await context.sync();
          //day1
          let d1 = { address: "E7", time: "08:00" }
          let d2 = { address: "F7", time: "08:30" }
          let d3 = { address: "G7", time: "09:00" }
          let d4 = { address: "H7", time: "09:30" }
          let d5 = { address: "I7", time: "10:00" }
          let d6 = { address: "J7", time: "10:30" }
          let d7 = { address: "K7", time: "11:00" }
          let d8 = { address: "L7", time: "11:30" }
          let d9 = { address: "M7", time: "12:00" }
          let d10 = { address: "N7", time: "12:30" }
          let d11 = { address: "O7", time: "13:00" }
          let d12 = { address: "P7", time: "13:30" }
          let d13 = { address: "Q7", time: "14:00" }
          let d14 = { address: "R7", time: "14:30" }
          let d15 = { address: "S7", time: "15:00" }
          let d16 = { address: "T7", time: "15:30" }
          let d17 = { address: "U7", time: "16:00" }
          let d18 = { address: "V7", time: "16:30" }
          let d19 = { address: "W7", time: "17:00" }
          let d20 = { address: "X7", time: "17:30" }
          let d21 = { address: "Y7", time: "18:00" }
          let d22 = { address: "Z7", time: "18:30" }
          let d23 = { address: "AA7", time: "19:00" }
          let d24 = { address: "AB7", time: "19:30" }
          let d25 = { address: "AC7", time: "20:00" }
          let d26 = { address: "AD7", time: "20:30" }
          let d27 = { address: "AE7", time: "21:00" }
          let d28 = { address: "AF7", time: "21:30" }
          let t1 = { address: "AG7", time: "22:00" }
          let t2 = { address: "AH7", time: "22:30" }
          let t3 = { address: "AI7", time: "23:00" }
          let t4 = { address: "AJ7", time: "23:30" }
          let timeallD1 = [d1, d2, d3, d4, d5, d6, d7, d8, d9, d10, d11, d12, d13, d14, d15, d16, d17, d18, d19, d20, d21,
            d22, d23, d24, d25, d26, d27, d28, t1, t2, t3, t4]
          //day2
          let t5 = { address: "AK7", time: "0:00" }
          let t6 = { address: "AL7", time: "0:30" }
          let t7 = { address: "AM7", time: "1:00" }
          let t8 = { address: "AN7", time: "1:30" }
          let t9 = { address: "AO7", time: "2:00" }
          let t10 = { address: "AP7", time: "2:30" }
          let t11 = { address: "AQ7", time: "3:00" }
          let t12 = { address: "AR7", time: "3:30" }
          let t13 = { address: "AS7", time: "4:00" }
          let t14 = { address: "AT7", time: "4:30" }
          let t15 = { address: "AU7", time: "5:00" }
          let t16 = { address: "AV7", time: "5:30" }
          let t17 = { address: "AW7", time: "6:00" }
          let t18 = { address: "AX7", time: "6:30" }
          let t19 = { address: "AY7", time: "7:00" }
          let t20 = { address: "AZ7", time: "7:30" }
          let t21 = { address: "BA7", time: "8:00" }
          let t22 = { address: "BB7", time: "8:30" }
          let t23 = { address: "BC7", time: "9:00" }
          let t24 = { address: "BD7", time: "9:30" }
          let t25 = { address: "BE7", time: "10:00" }
          let t26 = { address: "BF7", time: "10:30" }
          let t27 = { address: "BG7", time: "11:00" }
          let t28 = { address: "BH7", time: "11:30" }
          let t29 = { address: "BI7", time: "12:00" }
          let t30 = { address: "BJ7", time: "12:30" }
          let t31 = { address: "BK7", time: "13:00" }
          let t32 = { address: "BL7", time: "13:30" }
          let t33 = { address: "BM7", time: "14:00" }
          let t34 = { address: "BN7", time: "14:30" }
          let t35 = { address: "BO7", time: "15:00" }
          let t36 = { address: "BP7", time: "15:30" }
          let t37 = { address: "BQ7", time: "16:00" }
          let t38 = { address: "BR7", time: "16:30" }
          let t39 = { address: "BS7", time: "17:00" }
          let t40 = { address: "BT7", time: "17:30" }
          let t41 = { address: "BU7", time: "18:00" }
          let t42 = { address: "BV7", time: "18:30" }
          let t43 = { address: "BW7", time: "19:00" }
          let t44 = { address: "BX7", time: "19:30" }
          let t45 = { address: "BY7", time: "20:00" }
          let t46 = { address: "BZ7", time: "20:30" }
          let t47 = { address: "CA7", time: "21:00" }
          let t48 = { address: "CB7", time: "21:30" }
          let t49 = { address: "CC7", time: "22:00" }
          let t50 = { address: "CD7", time: "22:30" }
          let t51 = { address: "CE7", time: "23:00" }
          let t52 = { address: "CF7", time: "23:30" }
          let timeallD2 = [t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16, t17, t18, t19, t20,
            t21, t22, t23, t24, t25, t26, t27, t28, t29, t30, t31, t32, t33, t34, t35, t36, t37, t38, t39, t40,
            t41, t42, t43, t44, t45, t46, t47, t48, t49, t50, t51, t52]
          //day3
          let e1 = { address: "CG7", time: "0:00" }
          let e2 = { address: "CH7", time: "0:30" }
          let e3 = { address: "CI7", time: "1:00" }
          let e4 = { address: "CJ7", time: "1:30" }
          let e5 = { address: "CK7", time: "2:00" }
          let e6 = { address: "CL7", time: "2:30" }
          let e7 = { address: "CM7", time: "3:00" }
          let e8 = { address: "CN7", time: "3:30" }
          let e9 = { address: "CO7", time: "4:00" }
          let e10 = { address: "CP7", time: "4:30" }
          let e11 = { address: "CQ7", time: "5:00" }
          let e12 = { address: "CR7", time: "5:30" }
          let e13 = { address: "CS7", time: "6:00" }
          let e14 = { address: "CT7", time: "6:30" }
          let e15 = { address: "CU7", time: "7:00" }
          let e16 = { address: "CV7", time: "7:30" }
          let e17 = { address: "Cw7", time: "8:00" }
          let e18 = { address: "CX7", time: "8:30" }
          let e19 = { address: "CY7", time: "9:00" }
          let e20 = { address: "CZ7", time: "9:30" }
          let e21 = { address: "DA7", time: "10:00" }
          let e22 = { address: "DB7", time: "10:30" }
          let e23 = { address: "DC7", time: "11:00" }
          let e24 = { address: "DD7", time: "11:30" }
          let e25 = { address: "DE7", time: "12:00" }
          let e26 = { address: "DF7", time: "12:30" }
          let e27 = { address: "DG7", time: "13:00" }
          let e28 = { address: "DH7", time: "13:30" }
          let e29 = { address: "DI7", time: "14:00" }
          let e30 = { address: "DJ7", time: "14:30" }
          let e31 = { address: "DK7", time: "15:00" }
          let e32 = { address: "DL7", time: "15:30" }
          let e33 = { address: "DM7", time: "16:00" }
          let e34 = { address: "DN7", time: "16:30" }
          let e35 = { address: "DO7", time: "17:00" }
          let e36 = { address: "DP7", time: "17:30" }
          let e37 = { address: "DQ7", time: "18:00" }
          let e38 = { address: "DR7", time: "18:30" }
          let e39 = { address: "DS7", time: "19:00" }
          let e40 = { address: "DT7", time: "19:30" }
          let e41 = { address: "DU7", time: "20:00" }
          let e42 = { address: "DV7", time: "20:30" }
          let e43 = { address: "DW7", time: "21:00" }
          let e44 = { address: "DX7", time: "21:30" }
          let e45 = { address: "DY7", time: "22:00" }
          let e46 = { address: "DZ7", time: "22:30" }
          let e47 = { address: "EA7", time: "23:00" }
          let e48 = { address: "EB7", time: "23:30" }
          let timeallD3 = [e1, e2, e3, e4, e5, e6, e7, e8, e9, e10, e11, e12, e13, e14, e15, e16, e17, e18, e19, e20, e21, e22, e23, e24,
            e25, e26, e27, e28, e29, e30, e31, e32, e33, e34, e35, e36, e37, e38, e39, e40, e41, e42, e43, e44, e45, e46, e47, e48]
          // console.log(timeall)
          let taskinprogress = []
          let taskdone = []
          let taskdelay = []
          let taskIC = []
          let taskDC = []

          Executegroup(statusrangeg1.text, taskrangeg1.text, durationrangeg1.text, startrangeg1.text, endrangeg1.text, resourcerangeg1.text, dayg1, grouprangeg1.text, targetg1, statusrangeg1.address)
          if (activeg2 === "yes")
            Executegroup(statusrangeg2.text, taskrangeg2.text, durationrangeg2.text, startrangeg2.text, endrangeg2.text, resourcerangeg2.text, dayg2, grouprangeg2.text, targetg2, statusrangeg2.address)
          if (activeg3 === "yes")
            Executegroup(statusrangeg3.text, taskrangeg3.text, durationrangeg3.text, startrangeg3.text, endrangeg3.text, resourcerangeg3.text, dayg3, grouprangeg3.text, targetg3, statusrangeg3.address)
          if (activeg4 === "yes")
            Executegroup(statusrangeg4.text, taskrangeg4.text, durationrangeg4.text, startrangeg4.text, endrangeg4.text, resourcerangeg4.text, dayg4, grouprangeg4.text, targetg4, statusrangeg4.address)
          if (activeg5 === "yes")
            Executegroup(statusrangeg5.text, taskrangeg5.text, durationrangeg5.text, startrangeg5.text, endrangeg5.text, resourcerangeg5.text, dayg5, grouprangeg5.text, targetg5, statusrangeg5.address)
          if (activeg6 === "yes")
            Executegroup(statusrangeg6.text, taskrangeg6.text, durationrangeg6.text, startrangeg6.text, endrangeg6.text, resourcerangeg6.text, dayg6, grouprangeg6.text, targetg6, statusrangeg6.address)
          if (activeg7 === "yes")
            Executegroup(statusrangeg7.text, taskrangeg7.text, durationrangeg7.text, startrangeg7.text, endrangeg7.text, resourcerangeg7.text, dayg7, grouprangeg7.text, targetg7, statusrangeg7.address)
          if (activeg8 === "yes")
            Executegroup(statusrangeg8.text, taskrangeg8.text, durationrangeg8.text, startrangeg8.text, endrangeg8.text, resourcerangeg8.text, dayg8, grouprangeg8.text, targetg8, statusrangeg8.address)
          if (activeg9 === "yes")
            Executegroup(statusrangeg9.text, taskrangeg9.text, durationrangeg9.text, startrangeg9.text, endrangeg9.text, resourcerangeg9.text, dayg9, grouprangeg9.text, targetg9, statusrangeg9.address)
          if (activeg10 === "yes")
            Executegroup(statusrangeg10.text, taskrangeg10.text, durationrangeg10.text, startrangeg10.text, endrangeg10.text, resourcerangeg10.text, dayg10, grouprangeg10.text, targetg10, statusrangeg10.address)
          if (activeg11 === "yes")
            Executegroup(statusrangeg11.text, taskrangeg11.text, durationrangeg11.text, startrangeg11.text, endrangeg11.text, resourcerangeg11.text, dayg11, grouprangeg11.text, targetg11, statusrangeg11.address)
          if (activeg12 === "yes")
            Executegroup(statusrangeg12.text, taskrangeg12.text, durationrangeg12.text, startrangeg12.text, endrangeg12.text, resourcerangeg12.text, dayg12, grouprangeg12.text, targetg12, statusrangeg12.address)
          if (activeg13 === "yes")
            Executegroup(statusrangeg13.text, taskrangeg13.text, durationrangeg13.text, startrangeg13.text, endrangeg13.text, resourcerangeg13.text, dayg13, grouprangeg13.text, targetg13, statusrangeg13.address)
          if (activeg14 === "yes")
            Executegroup(statusrangeg14.text, taskrangeg14.text, durationrangeg14.text, startrangeg14.text, endrangeg14.text, resourcerangeg14.text, dayg14, grouprangeg14.text, targetg14, statusrangeg14.address)
          if (activeg15 === "yes")
            Executegroup(statusrangeg15.text, taskrangeg15.text, durationrangeg15.text, startrangeg15.text, endrangeg15.text, resourcerangeg15.text, dayg15, grouprangeg15.text, targetg15, statusrangeg15.address)
          if (activeg16 === "yes")
            Executegroup(statusrangeg16.text, taskrangeg16.text, durationrangeg16.text, startrangeg16.text, endrangeg16.text, resourcerangeg16.text, dayg16, grouprangeg16.text, targetg16, statusrangeg16.address)
          if (activeg17 === "yes")
            Executegroup(statusrangeg17.text, taskrangeg17.text, durationrangeg17.text, startrangeg17.text, endrangeg17.text, resourcerangeg17.text, dayg17, grouprangeg17.text, targetg17, statusrangeg17.address)
          if (activeg18 === "yes")
            Executegroup(statusrangeg18.text, taskrangeg18.text, durationrangeg18.text, startrangeg18.text, endrangeg18.text, resourcerangeg18.text, dayg18, grouprangeg18.text, targetg18, statusrangeg18.address)
          if (activeg19 === "yes")
            Executegroup(statusrangeg19.text, taskrangeg19.text, durationrangeg19.text, startrangeg19.text, endrangeg19.text, resourcerangeg19.text, dayg19, grouprangeg19.text, targetg19, statusrangeg19.address)
          if (activeg20 === "yes")
            Executegroup(statusrangeg20.text, taskrangeg20.text, durationrangeg20.text, startrangeg20.text, endrangeg20.text, resourcerangeg20.text, dayg20, grouprangeg20.text, targetg20, statusrangeg20.address)
          if (activeg21 === "yes")
            Executegroup(statusrangeg21.text, taskrangeg21.text, durationrangeg21.text, startrangeg21.text, endrangeg21.text, resourcerangeg21.text, dayg21, grouprangeg21.text, targetg21, statusrangeg21.address)
          if (activeg22 === "yes")
            Executegroup(statusrangeg22.text, taskrangeg22.text, durationrangeg22.text, startrangeg22.text, endrangeg22.text, resourcerangeg22.text, dayg22, grouprangeg22.text, targetg22, statusrangeg22.address)
          if (activeg23 === "yes")
            Executegroup(statusrangeg23.text, taskrangeg23.text, durationrangeg23.text, startrangeg23.text, endrangeg23.text, resourcerangeg23.text, dayg23, grouprangeg23.text, targetg23, statusrangeg23.address)
          if (activeg24 === "yes")
            Executegroup(statusrangeg24.text, taskrangeg24.text, durationrangeg24.text, startrangeg24.text, endrangeg24.text, resourcerangeg24.text, dayg24, grouprangeg24.text, targetg24, statusrangeg24.address)
          if (activeg25 === "yes")
            Executegroup(statusrangeg25.text, taskrangeg25.text, durationrangeg25.text, startrangeg25.text, endrangeg25.text, resourcerangeg25.text, dayg25, grouprangeg25.text, targetg25, statusrangeg25.address)
          if (activeg26 === "yes")
            Executegroup(statusrangeg26.text, taskrangeg26.text, durationrangeg26.text, startrangeg26.text, endrangeg26.text, resourcerangeg26.text, dayg26, grouprangeg26.text, targetg26, statusrangeg26.address)
          if (activeg27 === "yes")
            Executegroup(statusrangeg27.text, taskrangeg27.text, durationrangeg27.text, startrangeg27.text, endrangeg27.text, resourcerangeg27.text, dayg27, grouprangeg27.text, targetg27, statusrangeg27.address)
          if (activeg28 === "yes")
            Executegroup(statusrangeg28.text, taskrangeg28.text, durationrangeg28.text, startrangeg28.text, endrangeg28.text, resourcerangeg28.text, dayg28, grouprangeg28.text, targetg28, statusrangeg28.address)
          if (activeg29 === "yes")
            Executegroup(statusrangeg29.text, taskrangeg29.text, durationrangeg29.text, startrangeg29.text, endrangeg29.text, resourcerangeg29.text, dayg29, grouprangeg29.text, targetg29, statusrangeg29.address)
          if (activeg30 === "yes")
            Executegroup(statusrangeg30.text, taskrangeg30.text, durationrangeg30.text, startrangeg30.text, endrangeg30.text, resourcerangeg30.text, dayg30, grouprangeg30.text, targetg30, statusrangeg30.address)


          var inprogressdata = taskinprogress.toString();
          var donedata = taskdone.toString();
          var delaydata = taskdelay.toString();
          var ICdata = taskIC.toString()
          var CDdata = taskDC.toString()
          // console.log(inprogressdata.replace(/,/g, '\n'))
          // console.log(donedata.replace(/,/g, '\n'))
          // console.log(delaydata.replace(/,/g, '\n'))
          // console.log(ICdata.replace(/,/g, '\n'))
          // console.log(CDdata.replace(/,/g, '\n'))
          inprogresstask.values = inprogressdata.replace(/,/g, '\n');
          donetask.values = donedata.replace(/,/g, '\n')
          delaytask.values = delaydata.replace(/,/g, '\n')
          ICtask.values = ICdata.replace(/,/g, '\n')
          CDtask.values = CDdata.replace(/,/g, '\n')

          function Executegroup(xx, xy, xz, yx, yy, yz, zz, gg, row, statusadd) {
            //start G1
            // console.log(zz)
            let timedata = []
            if (zz === "1") {
              // console.log(timeallD1)
              timedata = timeallD1
            } else if (zz === "2") {
              // console.log(timeallD2)
              timedata = timeallD2
            } else if (zz === "3") {
              // console.log(timeallD3)
              timedata = timeallD3
            } else if (zz === "1,2") {
              let a = timeallD1.concat(timeallD2)
              timedata = a
              // console.log(a)
            } else if (zz === "2,3") {
              let a = timeallD2.concat(timeallD3)
              timedata = a
              // console.log(a)
            }
            // console.log(timedata)
            let G1 = xx;
            // console.log(G1)
            let r2 = checkdata(G1, "Not Start")
            let checkIC = checkdata(G1, "In-Complete")
            // console.log(r2)
            if (r2.length == G1.length) {
              // sendcolor.format.fill.color = "gray";
              inprogresstask.values = ""
              donetask.values = ""
              delaytask.values = ""
              ICtask.values = ""
              CDtask.values = ""
            } else if (checkIC.length < 1) {
              // console.log(yx)
              autofillstatus(statusadd, xx, yx, yy, xz)
              G1.map((data, index) => {
                // console.log(yx[index][0] + "/" + yy[index][0] + "/" + xz[index][0])
                if (data[0] == "In-Progress") {
                  // console.log(data + "/IP/" + index)
                  let a = xy[index][0] + "/" + xz[index][0] + "/" + yx[index][0] + "/" + yy[index][0] + "/" + yz[index][0]
                  if (inprogenable === "yes")
                    taskinprogress.push(a)
                  In_Progress(yx[index][0], yy[index][0], gg[index][0], timedata, row)
                } else if (data[0] == "Complete") {
                  // console.log(data + "/C/" + index)
                  let a = xy[index][0] + "/" + xz[index][0] + "/" + yx[index][0] + "/" + yy[index][0] + "/" + yz[index][0]
                  if (compenable === "yes")
                    taskdone.push(a)
                  // console.log(yx[index] + '/' + yy[index] + "/" + gg[index])
                  Complete(yx[index][0], yy[index][0], gg[index][0], timedata, row)
                } else if (data[0] == "Delay") {
                  // console.log(data + "/D/" + index)
                  let a = xy[index][0] + "/" + xz[index][0] + "/" + yx[index][0] + "/" + yy[index][0] + "/" + yz[index][0]
                  if (delayenable === "yes")
                    taskdelay.push(a)
                  Delay(yx[index][0], yy[index][0], gg[index][0], timedata, row)
                } else if (data[0] == "In-Complete") {
                  // console.log(data + "/IC/" + index)
                  let a = xy[index][0] + "/" + xz[index][0] + "/" + yx[index][0] + "/" + yy[index][0] + "/" + yz[index][0]
                  if (icompenable === "yes")
                    taskIC.push(a)
                  In_Complete(yx[index][0], yy[index][0], gg[index][0], timedata, row)
                } else if ((data[0] == "Complete-Delay")) {
                  // console.log(data + "/CD/" + index)
                  let a = xy[index][0] + "/" + xz[index][0] + "/" + yx[index][0] + "/" + yy[index][0] + "/" + yz[index][0]
                  if (cdelayenable === "yes")
                    taskDC.push(a)
                  Complete_Delay(yx[index][0], yy[index][0], gg[index][0], timedata, row)
                }
              })

            }
            //end G1
          }

          function checkdata(a, b) {
            let d = a.filter(data1 => {
              // console.log(data1)
              let dcheck = data1[0].includes(b)
              return dcheck
            })
            // console.log(d)
            return d
          }
          function autofillstatus(a, b, c, d, e) {
            let ts = Date.now();
            let aa = a.split("!")
            let ab = aa[1].split(":")
            let ac = processText(ab[0])
            let ad = ac[0]
            // console.log(ad)
            let r1 = parseInt(ad[1])

            function processText(inputText) {
              var output = [];
              var json = inputText.split(' ');
              json.forEach(function (item) {
                output.push(item.replace(/\'/g, '').split(/(\d+)/).filter(Boolean));
              });
              return output;
            }

            for (let i = 0; i < b.length; i++) {
              let r2 = r1 + i;
              let r3 = r2.toString()
              // console.log(c[i] + "/" + d[i] + "/" + e[1])
              let timestart = getDateFromHours(c[i][0]).getTime();
              let timeend = getDateFromHours(d[i][0]).getTime();
              // console.log(timestart + "/" + timeend)
              let du = e[i][0].split(":")
              let du1 = parseInt(du[0]) * 60 * 60 * 1000
              let du2 = parseInt(du[1]) * 60 * 1000
              let du3 = du1 + du2
              // console.log(a + "/" + ad[0] + r3)
              // console.log(timeend - timestart + "/" + du3)
              if (c[i][0] === "" && d[i][0] === '') {
                // console.log(ab[0] + r3)
                let s = sheet1.getRange(ad[0] + r3)
                s.load(["values", "address"]);
                s.values = "Not Start";
              } else if (c[i][0] !== null && d[i][0] === '') {
                // console.log(ab[0] + r3)
                let s = sheet1.getRange(ad[0] + r3)
                s.load(["values", "address"]);
                s.values = "In-Progress";
              } else if (c[i][0] !== null && d[i][0] !== null) {
                if ((timeend - timestart) <= du3) {
                  // console.log("complete")
                  let s = sheet1.getRange(ad[0] + r3)
                  s.load(["values", "address"]);
                  s.values = "Complete";
                } else if ((timeend - timestart) > du3) {
                  // console.log("complete")
                  let s = sheet1.getRange(ad[0] + r3)
                  s.load(["values", "address"]);
                  s.values = "Complete-Delay";
                } else if (c[i][0] !== null && d[i][0] === '' && (ts > timestart + du3)) {
                  let s = sheet1.getRange(ad[0] + r3)
                  s.load(["values", "address"]);
                  s.values = "Delay";
                }
              }
            }
          }
          function In_Progress(x, y, z, dd, r) {
            // console.log(x + "/" + y + "/" + dd.length)
            if (x !== null && y === '') {
              for (let ii = 0; ii < dd.length; ii++) {
                let timeplus = parseInt(getDateFromHours(dd[ii].time).getTime() + 1800000)
                // console.log(getDateFromHours(x).getTime() + "/" + timeplus)
                if (getDateFromHours(x).getTime() < timeplus && getDateFromHours(x).getTime() >= getDateFromHours(dd[ii].time).getTime()) {
                  // console.log("inprogress")
                  let j = (dd[ii].address)
                  let j2 = j.split('')
                  let j1 = j2[0] + j2[1] + r
                  let a1 = sheet2.getRange(j1)
                  // console.log(j1 + "/IP")
                  a1.load(["values", "address"]);
                  // a1.values = "IP";
                  a1.format.fill.color = "cyan";
                  // break;
                  // a1.format.fill.color = "blue";
                }
              }
            }
          }
          function Complete(x, y, z, dd, r) {
            // console.log(dd)
            if (x !== null && y !== null) {
              for (let ii = 0; ii < dd.length; ii++) {
                if (ii < dd.length - 1) {
                  if (((getDateFromHours(x).getTime() < getDateFromHours(dd[ii + 1].time).getTime() && parseInt(getDateFromHours(y).getTime()) > parseInt(getDateFromHours(dd[ii].time).getTime()))) || ((getDateFromHours(x).getTime() >= getDateFromHours(dd[ii].time).getTime() && getDateFromHours(y).getTime() < getDateFromHours(dd[ii + 1].time).getTime()))) {
                    // console.log(dd[ii].time)
                    let j = (dd[ii].address)
                    let j2 = j.split('')
                    let j1 = j2[0] + j2[1] + r
                    let a1 = sheet2.getRange(j1)
                    // a1.values = "C";
                    a1.format.fill.color = "green";

                  }
                }
              }
            }
          }
          function In_Complete(x, y, z, dd, r) {
            // console.log(x + "/" + y)
            if (x !== null && y !== null) {
              for (let ii = 0; ii < dd.length; ii++) {
                if (ii < dd.length - 1) {
                  if (((getDateFromHours(x).getTime() < getDateFromHours(dd[ii + 1].time).getTime() && parseInt(getDateFromHours(y).getTime()) > parseInt(getDateFromHours(dd[ii].time).getTime()))) || ((getDateFromHours(x).getTime() >= getDateFromHours(dd[ii].time).getTime() && getDateFromHours(y).getTime() < getDateFromHours(dd[ii + 1].time).getTime()))) {
                    console.log(dd[ii].time)
                    // let a;
                    // if (z == "1") {
                    //   a = targetg1
                    // }
                    let j = (dd[ii].address)
                    let j2 = j.split('')
                    let j1 = j2[0] + j2[1] + r
                    let a1 = sheet2.getRange(j1)
                    // a1.values = "IC";
                    a1.format.fill.color = "red";
                  }
                }
              }
            }
          }
          function Delay(x, y, z, dd, r) {
            // console.log(x + "/" + y)
            if (x !== null && y !== null) {
              for (let ii = 0; ii < dd.length; ii++) {
                if (ii < dd.length - 1) {
                  if (((getDateFromHours(x).getTime() < getDateFromHours(dd[ii + 1].time).getTime() && parseInt(getDateFromHours(y).getTime()) > parseInt(getDateFromHours(dd[ii].time).getTime()))) || ((getDateFromHours(x).getTime() >= getDateFromHours(dd[ii].time).getTime() && getDateFromHours(y).getTime() < getDateFromHours(dd[ii + 1].time).getTime()))) {
                    console.log(dd[ii].time)
                    // let a;
                    // if (z == "1") {
                    //   a = targetg1
                    // }
                    let j = (dd[ii].address)
                    let j2 = j.split('')
                    let j1 = j2[0] + j2[1] + r
                    let a1 = sheet2.getRange(j1)
                    // a1.values = "D";
                    a1.format.fill.color = "gray";
                  }
                }
              }
            }
          }
          function Complete_Delay(x, y, z, dd, r) {
            // console.log(x + "/" + y)
            if (x !== null && y !== null) {
              for (let ii = 0; ii < dd.length; ii++) {
                if (ii < dd.length - 1) {
                  if (((getDateFromHours(x).getTime() < getDateFromHours(dd[ii + 1].time).getTime() && parseInt(getDateFromHours(y).getTime()) > parseInt(getDateFromHours(dd[ii].time).getTime()))) || ((getDateFromHours(x).getTime() >= getDateFromHours(dd[ii].time).getTime() && getDateFromHours(y).getTime() < getDateFromHours(dd[ii + 1].time).getTime()))) {
                    // console.log(dd[ii].time)

                    let j = (dd[ii].address)
                    let j2 = j.split('')
                    let j1 = j2[0] + j2[1] + r
                    let a1 = sheet2.getRange(j1)
                    // a1.values = "CD";
                    a1.format.fill.color = "orange";
                  }
                }
              }
            }
          }
          function getDateFromHours(time) {
            // console.log(time)
            // console.log(typeof time)
            time = time.split(':');
            let now = new Date();
            return new Date(now.getFullYear(), now.getMonth(), now.getDate(), ...time);
          }
        });
      } catch (error) {
        console.error(error);
      }
      await context.sync();
    };
    // return () => clearInterval(interval);

  }

  //end dev
  render() {
    const { title, isOfficeInitialized } = this.props;
    const showtask = () => {
      let dashboardconfig = JSON.parse(localStorage.getItem('dashboard'));
      if (dashboardconfig === null) {
        dashboardconfig = localStorage.setItem("dashboard", JSON.stringify({
          sname: "",
          tname: "", inprogressaddress: "", completeaddress: "", delayaddress: "",
          Icompleteaddress: ""
        }))
      }
      let g1config = JSON.parse(localStorage.getItem('g1config'));
      if (g1config === null) {
        g1config = localStorage.setItem("g1config", JSON.stringify({
          statusg1: "",
          taskidg1: "", durationg1: "", startg1: "", endg1: "",
          resourceg1: "", groupg1: "", targetg1: "", dayg1: ""
        }))
      }
      let g2config = JSON.parse(localStorage.getItem('g2config'));
      if (g2config === null) {
        g2config = localStorage.setItem("g2config", JSON.stringify({
          statusg2: "",
          taskidg2: "", durationg2: "", startg2: "", endg2: "",
          resourceg2: "", groupg2: "", targetg2: "", dayg2: ""
        }))
      }
      console.log(g1config)
      return (
        <>
          <a><b>SETUP DASHBOARD</b></a>
          {/* <Stack horizontal tokens={stackTokens}> */}
          <TextField id="datasheet" label="DATA SHEET " required defaultValue={dashboardconfig.sname} />
          <TextField id="dashboard" label="DASHBOARD SHEET " required defaultValue={dashboardconfig.tname} />
          <select id="inprogressenable">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="inprogress" label="IN-PROGRESS STATUS " required defaultValue="" />
          <select id="completeenable">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="complete" label="COMPLETE STATUS " required defaultValue="" />
          <select id="incompleteenable">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="incomplete" label="IN-COMPLETE STATUS " required defaultValue="" />
          <select id="delayenable">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="delay" label="DELAY STATUS " required defaultValue="" />
          <select id="completedelayenable">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="completedelay" label="COMPLETE-DELAY STATUS " required defaultValue="" />
          {/* </Stack> */}
          {/* <hr></hr>
          <a><b>SETUP time by range of row</b></a>
          <TextField id="timerange" label="TIME RANGE " required defaultValue="AL7:GO7" /> */}
          <hr></hr>
          <a><b>SETUP Group 1</b></a>
          <TextField id="datastatusg1" label="Status COL G1 " required defaultValue={g1config.statusg1} />
          <TextField id="datataskidg1" label="TASK ID COL G1 " required defaultValue={g1config.taskidg1} />
          <TextField id="datadurationg1" label="DURATION COL G1 " required defaultValue={g1config.durationg1} />
          <TextField id="datastartg1" label="START COL G1 " required defaultValue={g1config.startg1} />
          <TextField id="dataendg1" label="END COL G1 " required defaultValue={g1config.endg1} />
          <TextField id="dataresourceg1" label="RESOURCE COL G1 " required defaultValue={g1config.resourceg1} />
          <TextField id="datagroupg1" label="GROUP COL G1 " required defaultValue={g1config.groupg1} />
          <TextField id="targetrowg1" label="Dashboard ROW G1 " required defaultValue={g1config.targetg1} />
          <TextField id="dayg1" label="DAY(1,2,3) G1 " required defaultValue={g1config.dayg1} />
          <hr></hr>
          <a><b>SETUP Group 2</b></a>
          <select id="activeg2">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg2" label="Status COL G2 " required defaultValue={g2config.statusg2} />
          <TextField id="datataskidg2" label="TASK ID COL G2 " required defaultValue={g2config.taskidg2} />
          <TextField id="datadurationg2" label="DURATION COL G2 " required defaultValue={g2config.durationg2} />
          <TextField id="datastartg2" label="START COL G2 " required defaultValue={g2config.startg2} />
          <TextField id="dataendg2" label="END COL G2 " required defaultValue={g2config.endg2} />
          <TextField id="dataresourceg2" label="RESOURCE COL G2 " required defaultValue={g2config.resourceg2} />
          <TextField id="datagroupg2" label="GROUP COL G2 " required defaultValue={g2config.groupg2} />
          <TextField id="targetrowg2" label="Dashboard ROW G2 " required defaultValue={g2config.targetg2} />
          <TextField id="dayg2" label="DAY(1,2,3) G2 " required defaultValue={g2config.dayg2} />
          <hr></hr>
          <a><b>SETUP Group 3</b></a>
          <select id="activeg3">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg3" label="Status COL G3 " required defaultValue="" />
          <TextField id="datataskidg3" label="TASK ID COL G3 " required defaultValue="" />
          <TextField id="datadurationg3" label="DURATION COL G3 " required defaultValue="" />
          <TextField id="datastartg3" label="START COL G3 " required defaultValue="" />
          <TextField id="dataendg3" label="END COL G3 " required defaultValue="" />
          <TextField id="dataresourceg3" label="RESOURCE COL G3 " required defaultValue="" />
          <TextField id="datagroupg3" label="GROUP COL G3 " required defaultValue="" />
          <TextField id="targetrowg3" label="Dashboard ROW G3 " required defaultValue="" />
          <TextField id="dayg3" label="DAY(1,2,3) G3 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 4</b></a>
          <select id="activeg4">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg4" label="Status COL G4 " required defaultValue="" />
          <TextField id="datataskidg4" label="TASK ID COL G4 " required defaultValue="" />
          <TextField id="datadurationg4" label="DURATION COL G4 " required defaultValue="" />
          <TextField id="datastartg4" label="START COL G4 " required defaultValue="" />
          <TextField id="dataendg4" label="END COL G4 " required defaultValue="" />
          <TextField id="dataresourceg4" label="RESOURCE COL G4 " required defaultValue="" />
          <TextField id="datagroupg4" label="GROUP COL G4 " required defaultValue="" />
          <TextField id="targetrowg4" label="Dashboard ROW G4 " required defaultValue="" />
          <TextField id="dayg4" label="DAY(1,2,3) G4 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 5</b></a>
          <select id="activeg5">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg5" label="Status COL G5 " required defaultValue="" />
          <TextField id="datataskidg5" label="TASK ID COL G5 " required defaultValue="" />
          <TextField id="datadurationg5" label="DURATION COL G5 " required defaultValue="" />
          <TextField id="datastartg5" label="START COL G5 " required defaultValue="" />
          <TextField id="dataendg5" label="END COL G5 " required defaultValue="" />
          <TextField id="dataresourceg5" label="RESOURCE COL G5 " required defaultValue="" />
          <TextField id="datagroupg5" label="GROUP COL G5 " required defaultValue="" />
          <TextField id="targetrowg5" label="Dashboard ROW G5 " required defaultValue="" />
          <TextField id="dayg5" label="DAY(1,2,3) G5 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 6</b></a>
          <select id="activeg6">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg6" label="Status COL G6 " required defaultValue="" />
          <TextField id="datataskidg6" label="TASK ID COL G6 " required defaultValue="" />
          <TextField id="datadurationg6" label="DURATION COL G6 " required defaultValue="" />
          <TextField id="datastartg6" label="START COL G6 " required defaultValue="" />
          <TextField id="dataendg6" label="END COL G6 " required defaultValue="" />
          <TextField id="dataresourceg6" label="RESOURCE COL G6 " required defaultValue="" />
          <TextField id="datagroupg6" label="GROUP COL G6 " required defaultValue="" />
          <TextField id="targetrowg6" label="Dashboard ROW G6 " required defaultValue="" />
          <TextField id="dayg6" label="DAY(1,2,3) G6 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 7</b></a>
          <select id="activeg7">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg7" label="Status COL G7 " required defaultValue="" />
          <TextField id="datataskidg7" label="TASK ID COL G7 " required defaultValue="" />
          <TextField id="datadurationg7" label="DURATION COL G7 " required defaultValue="" />
          <TextField id="datastartg7" label="START COL G7 " required defaultValue="" />
          <TextField id="dataendg7" label="END COL G7 " required defaultValue="" />
          <TextField id="dataresourceg7" label="RESOURCE COL G7 " required defaultValue="" />
          <TextField id="datagroupg7" label="GROUP COL G7 " required defaultValue="" />
          <TextField id="targetrowg7" label="Dashboard ROW G7 " required defaultValue="" />
          <TextField id="dayg7" label="DAY(1,2,3) G7 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 8</b></a>
          <select id="activeg8">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg8" label="Status COL G8 " required defaultValue="" />
          <TextField id="datataskidg8" label="TASK ID COL G8 " required defaultValue="" />
          <TextField id="datadurationg8" label="DURATION COL G8 " required defaultValue="" />
          <TextField id="datastartg8" label="START COL G8 " required defaultValue="" />
          <TextField id="dataendg8" label="END COL G8 " required defaultValue="" />
          <TextField id="dataresourceg8" label="RESOURCE COL G8 " required defaultValue="" />
          <TextField id="datagroupg8" label="GROUP COL G8 " required defaultValue="" />
          <TextField id="targetrowg8" label="Dashboard ROW G8 " required defaultValue="" />
          <TextField id="dayg8" label="DAY(1,2,3) G8 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 9</b></a>
          <select id="activeg9">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg9" label="Status COL G9 " required defaultValue="" />
          <TextField id="datataskidg9" label="TASK ID COL G9 " required defaultValue="" />
          <TextField id="datadurationg9" label="DURATION COL G9 " required defaultValue="" />
          <TextField id="datastartg9" label="START COL G9 " required defaultValue="" />
          <TextField id="dataendg9" label="END COL G9 " required defaultValue="" />
          <TextField id="dataresourceg9" label="RESOURCE COL G9 " required defaultValue="" />
          <TextField id="datagroupg9" label="GROUP COL G9 " required defaultValue="" />
          <TextField id="targetrowg9" label="Dashboard ROW G9 " required defaultValue="" />
          <TextField id="dayg9" label="DAY(1,2,3) G9 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 10</b></a>
          <select id="activeg10">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg10" label="Status COL G10 " required defaultValue="" />
          <TextField id="datataskidg10" label="TASK ID COL G10 " required defaultValue="" />
          <TextField id="datadurationg10" label="DURATION COL G10 " required defaultValue="" />
          <TextField id="datastartg10" label="START COL G10 " required defaultValue="" />
          <TextField id="dataendg10" label="END COL G10 " required defaultValue="" />
          <TextField id="dataresourceg10" label="RESOURCE COL G10 " required defaultValue="" />
          <TextField id="datagroupg10" label="GROUP COL G10 " required defaultValue="" />
          <TextField id="targetrowg10" label="Dashboard ROW G10 " required defaultValue="" />
          <TextField id="dayg10" label="DAY(1,2,3) G10 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 11</b></a>
          <select id="activeg11">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg11" label="Status COL G11 " required defaultValue="" />
          <TextField id="datataskidg11" label="TASK ID COL G11 " required defaultValue="" />
          <TextField id="datadurationg11" label="DURATION COL G11 " required defaultValue="" />
          <TextField id="datastartg11" label="START COL G11 " required defaultValue="" />
          <TextField id="dataendg11" label="END COL G11 " required defaultValue="" />
          <TextField id="dataresourceg11" label="RESOURCE COL G11 " required defaultValue="" />
          <TextField id="datagroupg11" label="GROUP COL G11 " required defaultValue="" />
          <TextField id="targetrowg11" label="Dashboard ROW G11 " required defaultValue="" />
          <TextField id="dayg11" label="DAY(1,2,3) G11 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 12</b></a>
          <select id="activeg12">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg12" label="Status COL G12 " required defaultValue="" />
          <TextField id="datataskidg12" label="TASK ID COL G12 " required defaultValue="" />
          <TextField id="datadurationg12" label="DURATION COL G12 " required defaultValue="" />
          <TextField id="datastartg12" label="START COL G12 " required defaultValue="" />
          <TextField id="dataendg12" label="END COL G12 " required defaultValue="" />
          <TextField id="dataresourceg12" label="RESOURCE COL G12 " required defaultValue="" />
          <TextField id="datagroupg12" label="GROUP COL G12 " required defaultValue="" />
          <TextField id="targetrowg12" label="Dashboard ROW G12 " required defaultValue="" />
          <TextField id="dayg12" label="DAY(1,2,3) G12 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 13</b></a>
          <select id="activeg13">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg13" label="Status COL G13 " required defaultValue="" />
          <TextField id="datataskidg13" label="TASK ID COL G13 " required defaultValue="" />
          <TextField id="datadurationg13" label="DURATION COL G13 " required defaultValue="" />
          <TextField id="datastartg13" label="START COL G13 " required defaultValue="" />
          <TextField id="dataendg13" label="END COL G13 " required defaultValue="" />
          <TextField id="dataresourceg13" label="RESOURCE COL G13 " required defaultValue="" />
          <TextField id="datagroupg13" label="GROUP COL G13 " required defaultValue="" />
          <TextField id="targetrowg13" label="Dashboard ROW G13 " required defaultValue="" />
          <TextField id="dayg13" label="DAY(1,2,3) G13 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 14</b></a>
          <select id="activeg14">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg14" label="Status COL G14 " required defaultValue="" />
          <TextField id="datataskidg14" label="TASK ID COL G14 " required defaultValue="" />
          <TextField id="datadurationg14" label="DURATION COL G14 " required defaultValue="" />
          <TextField id="datastartg14" label="START COL G14 " required defaultValue="" />
          <TextField id="dataendg14" label="END COL G14 " required defaultValue="" />
          <TextField id="dataresourceg14" label="RESOURCE COL G14 " required defaultValue="" />
          <TextField id="datagroupg14" label="GROUP COL G14 " required defaultValue="" />
          <TextField id="targetrowg14" label="Dashboard ROW G14 " required defaultValue="" />
          <TextField id="dayg14" label="DAY(1,2,3) G14 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 15</b></a>
          <select id="activeg15">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg15" label="Status COL G15 " required defaultValue="" />
          <TextField id="datataskidg15" label="TASK ID COL G15 " required defaultValue="" />
          <TextField id="datadurationg15" label="DURATION COL G15 " required defaultValue="" />
          <TextField id="datastartg15" label="START COL G15 " required defaultValue="" />
          <TextField id="dataendg15" label="END COL G15 " required defaultValue="" />
          <TextField id="dataresourceg15" label="RESOURCE COL G15 " required defaultValue="" />
          <TextField id="datagroupg15" label="GROUP COL G15 " required defaultValue="" />
          <TextField id="targetrowg15" label="Dashboard ROW G15 " required defaultValue="" />
          <TextField id="dayg15" label="DAY(1,2,3) G15 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 16</b></a>
          <select id="activeg16">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg16" label="Status COL G16 " required defaultValue="" />
          <TextField id="datataskidg16" label="TASK ID COL G16 " required defaultValue="" />
          <TextField id="datadurationg16" label="DURATION COL G16 " required defaultValue="" />
          <TextField id="datastartg16" label="START COL G16 " required defaultValue="" />
          <TextField id="dataendg16" label="END COL G16 " required defaultValue="" />
          <TextField id="dataresourceg16" label="RESOURCE COL G16 " required defaultValue="" />
          <TextField id="datagroupg16" label="GROUP COL G16 " required defaultValue="" />
          <TextField id="targetrowg16" label="Dashboard ROW G16 " required defaultValue="" />
          <TextField id="dayg16" label="DAY(1,2,3) G16 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 17</b></a>
          <select id="activeg17">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg17" label="Status COL G17 " required defaultValue="" />
          <TextField id="datataskidg17" label="TASK ID COL G17 " required defaultValue="" />
          <TextField id="datadurationg17" label="DURATION COL G17 " required defaultValue="" />
          <TextField id="datastartg17" label="START COL G17 " required defaultValue="" />
          <TextField id="dataendg17" label="END COL G17 " required defaultValue="" />
          <TextField id="dataresourceg17" label="RESOURCE COL G17 " required defaultValue="" />
          <TextField id="datagroupg17" label="GROUP COL G17 " required defaultValue="" />
          <TextField id="targetrowg17" label="Dashboard ROW G17 " required defaultValue="" />
          <TextField id="dayg17" label="DAY(1,2,3) G17 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 18</b></a>
          <select id="activeg18">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg18" label="Status COL G18 " required defaultValue="" />
          <TextField id="datataskidg18" label="TASK ID COL G18 " required defaultValue="" />
          <TextField id="datadurationg18" label="DURATION COL G18 " required defaultValue="" />
          <TextField id="datastartg18" label="START COL G18 " required defaultValue="" />
          <TextField id="dataendg18" label="END COL G18 " required defaultValue="" />
          <TextField id="dataresourceg18" label="RESOURCE COL G18 " required defaultValue="" />
          <TextField id="datagroupg18" label="GROUP COL G18 " required defaultValue="" />
          <TextField id="targetrowg18" label="Dashboard ROW G18 " required defaultValue="" />
          <TextField id="dayg18" label="DAY(1,2,3) G18 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 19</b></a>
          <select id="activeg19">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg19" label="Status COL G19 " required defaultValue="" />
          <TextField id="datataskidg19" label="TASK ID COL G19 " required defaultValue="" />
          <TextField id="datadurationg19" label="DURATION COL G19 " required defaultValue="" />
          <TextField id="datastartg19" label="START COL G19 " required defaultValue="" />
          <TextField id="dataendg19" label="END COL G19 " required defaultValue="" />
          <TextField id="dataresourceg19" label="RESOURCE COL G19 " required defaultValue="" />
          <TextField id="datagroupg19" label="GROUP COL G19 " required defaultValue="" />
          <TextField id="targetrowg19" label="Dashboard ROW G19 " required defaultValue="" />
          <TextField id="dayg19" label="DAY(1,2,3) G19 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 20</b></a>
          <select id="activeg20">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg20" label="Status COL G20 " required defaultValue="" />
          <TextField id="datataskidg20" label="TASK ID COL G20 " required defaultValue="" />
          <TextField id="datadurationg20" label="DURATION COL G20 " required defaultValue="" />
          <TextField id="datastartg20" label="START COL G20 " required defaultValue="" />
          <TextField id="dataendg20" label="END COL G20 " required defaultValue="" />
          <TextField id="dataresourceg20" label="RESOURCE COL G20 " required defaultValue="" />
          <TextField id="datagroupg20" label="GROUP COL G20 " required defaultValue="" />
          <TextField id="targetrowg20" label="Dashboard ROW G20 " required defaultValue="" />
          <TextField id="dayg20" label="DAY(1,2,3) G20 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 21</b></a>
          <select id="activeg21">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg21" label="Status COL G21 " required defaultValue="" />
          <TextField id="datataskidg21" label="TASK ID COL G21 " required defaultValue="" />
          <TextField id="datadurationg21" label="DURATION COL G21 " required defaultValue="" />
          <TextField id="datastartg21" label="START COL G21 " required defaultValue="" />
          <TextField id="dataendg21" label="END COL G21 " required defaultValue="" />
          <TextField id="dataresourceg21" label="RESOURCE COL G21 " required defaultValue="" />
          <TextField id="datagroupg21" label="GROUP COL G21 " required defaultValue="" />
          <TextField id="targetrowg21" label="Dashboard ROW G21 " required defaultValue="" />
          <TextField id="dayg21" label="DAY(1,2,3) G21 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 22</b></a>
          <select id="activeg22">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg22" label="Status COL G22 " required defaultValue="" />
          <TextField id="datataskidg22" label="TASK ID COL G22 " required defaultValue="" />
          <TextField id="datadurationg22" label="DURATION COL G22 " required defaultValue="" />
          <TextField id="datastartg22" label="START COL G22 " required defaultValue="" />
          <TextField id="dataendg22" label="END COL G22 " required defaultValue="" />
          <TextField id="dataresourceg22" label="RESOURCE COL G22 " required defaultValue="" />
          <TextField id="datagroupg22" label="GROUP COL G22 " required defaultValue="" />
          <TextField id="targetrowg22" label="Dashboard ROW G22 " required defaultValue="" />
          <TextField id="dayg22" label="DAY(1,2,3) G22 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 23</b></a>
          <select id="activeg23">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg23" label="Status COL G23 " required defaultValue="" />
          <TextField id="datataskidg23" label="TASK ID COL G23 " required defaultValue="" />
          <TextField id="datadurationg23" label="DURATION COL G23 " required defaultValue="" />
          <TextField id="datastartg23" label="START COL G23 " required defaultValue="" />
          <TextField id="dataendg23" label="END COL G23 " required defaultValue="" />
          <TextField id="dataresourceg23" label="RESOURCE COL G23 " required defaultValue="" />
          <TextField id="datagroupg23" label="GROUP COL G23 " required defaultValue="" />
          <TextField id="targetrowg23" label="Dashboard ROW G23 " required defaultValue="" />
          <TextField id="dayg23" label="DAY(1,2,3) G23 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 24</b></a>
          <select id="activeg24">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg24" label="Status COL G24 " required defaultValue="" />
          <TextField id="datataskidg24" label="TASK ID COL G24 " required defaultValue="" />
          <TextField id="datadurationg24" label="DURATION COL G24 " required defaultValue="" />
          <TextField id="datastartg24" label="START COL G24 " required defaultValue="" />
          <TextField id="dataendg24" label="END COL G24 " required defaultValue="" />
          <TextField id="dataresourceg24" label="RESOURCE COL G24 " required defaultValue="" />
          <TextField id="datagroupg24" label="GROUP COL G24 " required defaultValue="" />
          <TextField id="targetrowg24" label="Dashboard ROW G24 " required defaultValue="" />
          <TextField id="dayg24" label="DAY(1,2,3) G24 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 25</b></a>
          <select id="activeg25">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg25" label="Status COL G25 " required defaultValue="" />
          <TextField id="datataskidg25" label="TASK ID COL G25 " required defaultValue="" />
          <TextField id="datadurationg25" label="DURATION COL G25 " required defaultValue="" />
          <TextField id="datastartg25" label="START COL G25 " required defaultValue="" />
          <TextField id="dataendg25" label="END COL G25 " required defaultValue="" />
          <TextField id="dataresourceg25" label="RESOURCE COL G25 " required defaultValue="" />
          <TextField id="datagroupg25" label="GROUP COL G25 " required defaultValue="" />
          <TextField id="targetrowg25" label="Dashboard ROW G25 " required defaultValue="" />
          <TextField id="dayg25" label="DAY(1,2,3) G25 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 26</b></a>
          <select id="activeg26">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg26" label="Status COL G26 " required defaultValue="" />
          <TextField id="datataskidg26" label="TASK ID COL G26 " required defaultValue="" />
          <TextField id="datadurationg26" label="DURATION COL G26 " required defaultValue="" />
          <TextField id="datastartg26" label="START COL G26 " required defaultValue="" />
          <TextField id="dataendg26" label="END COL G26 " required defaultValue="" />
          <TextField id="dataresourceg26" label="RESOURCE COL G26 " required defaultValue="" />
          <TextField id="datagroupg26" label="GROUP COL G26 " required defaultValue="" />
          <TextField id="targetrowg26" label="Dashboard ROW G26 " required defaultValue="" />
          <TextField id="dayg26" label="DAY(1,2,3) G26 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 27</b></a>
          <select id="activeg27">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg27" label="Status COL G27 " required defaultValue="" />
          <TextField id="datataskidg27" label="TASK ID COL G27 " required defaultValue="" />
          <TextField id="datadurationg27" label="DURATION COL G27 " required defaultValue="" />
          <TextField id="datastartg27" label="START COL G27 " required defaultValue="" />
          <TextField id="dataendg27" label="END COL G27 " required defaultValue="" />
          <TextField id="dataresourceg27" label="RESOURCE COL G27 " required defaultValue="" />
          <TextField id="datagroupg27" label="GROUP COL G27 " required defaultValue="" />
          <TextField id="targetrowg27" label="Dashboard ROW G27 " required defaultValue="" />
          <TextField id="dayg27" label="DAY(1,2,3) G27 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 28</b></a>
          <select id="activeg28">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg28" label="Status COL G28 " required defaultValue="" />
          <TextField id="datataskidg28" label="TASK ID COL G28 " required defaultValue="" />
          <TextField id="datadurationg28" label="DURATION COL G28 " required defaultValue="" />
          <TextField id="datastartg28" label="START COL G28 " required defaultValue="" />
          <TextField id="dataendg28" label="END COL G28 " required defaultValue="" />
          <TextField id="dataresourceg28" label="RESOURCE COL G28 " required defaultValue="" />
          <TextField id="datagroupg28" label="GROUP COL G28 " required defaultValue="" />
          <TextField id="targetrowg28" label="Dashboard ROW G28 " required defaultValue="" />
          <TextField id="dayg28" label="DAY(1,2,3) G28 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 29</b></a>
          <select id="activeg29">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg29" label="Status COL G29 " required defaultValue="" />
          <TextField id="datataskidg29" label="TASK ID COL G29 " required defaultValue="" />
          <TextField id="datadurationg29" label="DURATION COL G29 " required defaultValue="" />
          <TextField id="datastartg29" label="START COL G29 " required defaultValue="" />
          <TextField id="dataendg29" label="END COL G29 " required defaultValue="" />
          <TextField id="dataresourceg29" label="RESOURCE COL G29 " required defaultValue="" />
          <TextField id="datagroupg29" label="GROUP COL G29 " required defaultValue="" />
          <TextField id="targetrowg29" label="Dashboard ROW G29 " required defaultValue="" />
          <TextField id="dayg29" label="DAY(1,2,3) G29 " required defaultValue="2" />
          <hr></hr>
          <a><b>SETUP Group 30</b></a>
          <select id="activeg30">
            <option value="no" selected>no</option>
            <option value="yes">yes</option>
          </select>
          <TextField id="datastatusg30" label="Status COL G30 " required defaultValue="" />
          <TextField id="datataskidg30" label="TASK ID COL G30 " required defaultValue="" />
          <TextField id="datadurationg30" label="DURATION COL G30 " required defaultValue="" />
          <TextField id="datastartg30" label="START COL G30 " required defaultValue="" />
          <TextField id="dataendg30" label="END COL G30 " required defaultValue="" />
          <TextField id="dataresourceg30" label="RESOURCE COL G30 " required defaultValue="" />
          <TextField id="datagroupg30" label="GROUP COL G30 " required defaultValue="" />
          <TextField id="targetrowg30" label="Dashboard ROW G30 " required defaultValue="" />
          <TextField id="dayg30" label="DAY(1,2,3) G30 " required defaultValue="2" />
        </>
      )
    }
    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <>

        <div className="ms-welcome">
          {/* <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" /> */}
          <HeroList message="BAY DC RE-LOCATION" items={this.state.listItems}>
            {/* <p className="ms-font-l">              
            </p> */}

            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={() => this.active(1)}>
              RUN
            </DefaultButton><span>{this.state.check ? "running" : "not running"}</span>
            <hr></hr>
            {showtask()}


            {/* <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={() => this.active(0)}>
            STOP
          </DefaultButton> */}
          </HeroList>
        </div>
      </>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
