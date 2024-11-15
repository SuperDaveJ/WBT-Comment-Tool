//For more functions, see NOAA series of surveys.
function answeredAll(nQs) {
	//check to see if all questions are rated.  Radio button are required but comments are optional.
	var nGrade = 4;
	//check all radio buttons
	nUnanswered = 0;
  	for (i=0; i<nQs; i++) {
		//for all radio buttons
		rButton = eval("document.forms['form1'].radio" + i);
		if ( !(rButton[0].checked) && !(rButton[1].checked) && !(rButton[2].checked) && !(rButton[3].checked) ) 	{
				nUnanswered += 1;
		}
  	}
	if (nUnanswered > 0) {
		alert ("You have " + nUnanswered + " questions not rated. \rPlease answer all questions.");
		return false;
	}
	return true;
}

function disableAll(nQs) {
	//This function works but not used.
	var nGrade = 4;
	//Disable all radio buttons
  	for (i=0; i<nQs; i++) {
		//for relevance buttons
		rButton = eval("document.forms['form1'].radio" + i);
		for (k=0; k<nGrade; k++) {
			rButton[k].disabled = true;
		}
  	}
}
