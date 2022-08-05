// JavaScript source code
var Winnable_ProjectDetails_BAL = {
	initializePage: function () {
		alert("This is from External Project Details");
		
		initializeMetronicScriptBundle();
		renderPageControls();
	}
};
function renderPageControls() {
	console.log("test");
	
}

function initializeMetronicScriptBundle()
{
	
// Boostrap & 3rd-party components initialization
KTApp.init();

// Boostrap & 3rd-party components dynamic creation
KTApp.createInstances();
	
KTMenu.createInstances();
KTDrawer.createInstances();
KTScroll.createInstances();
KTScrolltop.createInstances();
KTSticky.createInstances();
KTDialer.createInstances();
KTImageInput.createInstances();
KTPasswordMeter.createInstances();
KTSwapper.createInstances();
KTToggle.createInstances();

}
