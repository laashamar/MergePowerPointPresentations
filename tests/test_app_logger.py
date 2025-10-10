import logging

# MODIFIED: Import from the package
from merge_powerpoint.app_logger import setup_logging

def test_setup_logging_configures_root_logger():
    """
    Test that calling setup_logging configures the root logger correctly.
    """
    # Reset logging to a clean state before the test
    logging.getLogger().handlers = []
    
    setup_logging()
    
    logger = logging.getLogger()
    
    # Check that at least one handler has been added
    assert len(logger.handlers) > 0
    # Check that the level is set to INFO
    assert logger.level == logging.INFO

