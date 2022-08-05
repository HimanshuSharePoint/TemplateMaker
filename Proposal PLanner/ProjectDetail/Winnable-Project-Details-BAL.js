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
	
KTMenu.init();
KTMenu.createInstances();

KTDrawer.init();
KTDrawer.createInstances();

KTScroll.init();
KTScroll.createInstances();

KTScrolltop.init();
KTScrolltop.createInstances();

KTSticky.init();
KTSticky.createInstances();

KTDialer.init();
KTDialer.createInstances();

KTImageInput.init();
KTImageInput.createInstances();

KTPasswordMeter.init();
KTPasswordMeter.createInstances();

KTSwapper.init();
KTSwapper.createInstances();

KTToggle.init();
KTToggle.createInstances();

}
