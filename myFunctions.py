def clearSendkey(driver, input1, input2):
    driver.find_element_by_id(input1).clear()
    driver.find_element_by_id(input1).send_keys(input2)

def clearSendkey0(driver, input1, input2):
    driver.find_element_by_link_text(input1).clear()
    driver.find_element_by_link_text(input1).send_keys(input2)
